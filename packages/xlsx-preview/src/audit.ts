/**
 * The audit pass: finding numbers a model *states* that its own formulas do not
 * support.
 *
 * There are two shapes of this, and only one of them needs the engine.
 *
 * **Stale cache.** A formula cell whose stored `<v>` disagrees with what the
 * formula computes. The engine reports these directly, because it keeps both
 * values instead of overwriting one with the other.
 *
 * **Hardcoded cell.** A literal number sitting where its neighbours hold
 * formulas — `=SUM(B4:B5)`, `=SUM(C4:C5)`, `=SUM(D4:D5)`, then `1261.2`. This is
 * the characteristic failure of a generated model: the writer computed one
 * figure in its head and typed it in. There is no formula to disagree with, so
 * no value comparison can find it; it is only visible in the *shape* of the
 * neighbourhood.
 *
 * The detector infers what the cell would have held by translating a neighbour's
 * formula to its position and evaluating it — so the finding is specific
 * ("D6 states 1261.2; SUM(D4:D5) is 1221.2") rather than a vague suspicion. It
 * requires at least two neighbours sharing one shape, because a single
 * neighbouring formula is a pattern of one and would fire on every summary row.
 *
 * ## "Neighbours" has to mean neighbours
 *
 * The first implementation took *every* formula cell in the literal's row or
 * column as a candidate neighbour, with no locality requirement of any kind, and
 * then extrapolated a shape shared by any two of them onto any literal on the
 * same line. Against `apps/demo/public/financial-model-nocache.xlsx` — a file in
 * this repository — that produced **33 findings, every one of them false**.
 *
 * `B6` (`=B4+B5`) and `B18` (`=B16+B17`) sit twelve rows apart and share one
 * shape, so "each cell is the sum of the two above it" was generalised down the
 * entire column and then applied *upward* onto rows 3–5 — which are the model's
 * hand-entered inputs and a header row of date serials. The detector announced
 * that `B3` (46022, a date) should have held `B1+B2`, which is 0, because `B1`
 * and `B2` are the title rows and empty.
 *
 * A false finding here is worse than silence. It is specific, confident, and it
 * names a cell and a number — so acting on one means editing a model that was
 * correct.
 *
 * Three rules now stand between a shape and a finding. A shape must be an actual
 * *band* (`isLocalPattern`): its members occupy at least half of their own span,
 * nothing foreign sits inside that span, and the literal is inside the band or
 * immediately after it — never immediately before it, because the cell in front
 * of a run of formulas is where models seed a series rather than where they get
 * one wrong. And the literal must belong to the same *series* as the band
 * (`sameSeries`): a formula that would read raw inputs, sitting beside formulas
 * that read computed ones, is the first cell of another column of data, not a
 * missing copy of its neighbours.
 *
 * Measured on the ten-workbook real corpus (`eval/real`), which is real business
 * models rather than anything written for this test: **900+ findings before,
 * capped at 200 on four of the ten files, and 23 after** — roughly two per
 * workbook, and the survivors have the right shape (a total that misses its own
 * column sum by 167k; a stated 6731.67 where the formula gives 6731.69; round
 * -20000 plugs where an average computes -16849).
 *
 * That is a large improvement and it is **not a precision guarantee**. Nobody has
 * ground truth for those 23. A second, independent corpus of machine-generated
 * models was measured separately and every finding there was judged false, for a
 * reason worth carrying: a generated model routinely stacks a linked-data block
 * above a typed-judgment block in one column, and treating a column as one
 * homogeneous thing is the assumption behind this whole detector. Treat the
 * output as a place to look, not as a defect list.
 */

import { unparse, walk, Workbook, type Node, type Scalar } from '@xlscalc/formula-engine';
import { plainText } from './format.js';

export interface HardcodedFinding {
  sheet: number;
  sheetName: string;
  address: string;
  row: number;
  col: number;
  /** The literal the file states. */
  stated: Scalar;
  /** What the inferred formula computes instead. */
  expected: Scalar;
  /** The formula its neighbours use, translated to this cell. */
  formula: string;
  axis: 'row' | 'column';
  /** Addresses of the neighbours that established the pattern. */
  pattern: string[];
}

interface CellRef {
  row: number;
  col: number;
}

const EPSILON = 1e-9;

export function findHardcoded(
  engine: Workbook,
  literals: Array<{ sheet: number; row: number; col: number; value: Scalar }>,
  limit = 200,
): HardcodedFinding[] {
  const out: HardcodedFinding[] = [];

  // Index the formula cells by sheet, so the neighbourhood scan is local.
  const bySheet = new Map<number, { rows: Map<number, CellRef[]>; cols: Map<number, CellRef[]> }>();
  engine.sheets.forEach((sheet) => {
    const rows = new Map<number, CellRef[]>();
    const cols = new Map<number, CellRef[]>();
    for (const key of sheet.formulas.keys()) {
      const row = Math.floor(key / 16384);
      const col = key - row * 16384;
      push(rows, row, { row, col });
      push(cols, col, { row, col });
    }
    bySheet.set(sheet.index, { rows, cols });
  });

  // And index the numeric literals the same way. A band of `=SUM(...)` cells
  // interrupted by the very cell we are judging is still a band, and so is one
  // interrupted by a second hardcoded number beside it — both are common, and
  // neither should cost us the finding. Anything else inside the span (an empty
  // cell, a label, a formula of a different shape) means the run is not one
  // neighbourhood and the pattern must not be extended across it.
  const literalsBySheet = new Map<number, { rows: Map<number, Set<number>>; cols: Map<number, Set<number>> }>();
  for (const literal of literals) {
    if (typeof literal.value !== 'number') continue;
    let index = literalsBySheet.get(literal.sheet);
    if (!index) {
      index = { rows: new Map(), cols: new Map() };
      literalsBySheet.set(literal.sheet, index);
    }
    add(index.rows, literal.row, literal.col);
    add(index.cols, literal.col, literal.row);
  }

  for (const literal of literals) {
    if (out.length >= limit) break;
    if (typeof literal.value !== 'number') continue;

    const index = bySheet.get(literal.sheet);
    if (!index) continue;
    const lits = literalsBySheet.get(literal.sheet);

    const finding =
      inferFrom(engine, literal, index.rows.get(literal.row) ?? [], 'row', lits?.rows.get(literal.row)) ??
      inferFrom(engine, literal, index.cols.get(literal.col) ?? [], 'column', lits?.cols.get(literal.col));
    if (finding) out.push(finding);
  }
  return out;
}

function add<K>(map: Map<K, Set<number>>, key: K, value: number): void {
  const set = map.get(key);
  if (set) set.add(value);
  else map.set(key, new Set([value]));
}

/**
 * Members must occupy at least this fraction of their own span — expressed as a
 * divisor, so 2 means "at least half". Two formula cells at opposite ends of a
 * long column are not a pattern, however identical their shape.
 */
const MIN_DENSITY_DIVISOR = 2;

/**
 * Is this shape a genuine local band, with the literal in or beside it?
 *
 * Three conditions, each of which kills a distinct way the old rule went wrong:
 *
 *   1. **Density** — the members occupy at least half the span they cover, so
 *      `B6` and `B18` with eleven unrelated rows between them stop being a
 *      "pattern". This runs first because it also bounds everything below: the
 *      gap scan can never be longer than twice the member count.
 *   2. **Adjacency** — the literal sits inside the band or immediately beside
 *      it. A total at the end of a row of subtotals qualifies; a cell three rows
 *      above the nearest formula does not.
 *   3. **Nothing foreign inside** — every position within the span is a member
 *      of this shape, the literal itself, or another number (which may be a
 *      second hardcoded cell, judged on its own turn). A label, an empty cell or
 *      a formula of some other shape means these are two neighbourhoods, not
 *      one, and a rule from one of them must not be extended over the other.
 */
function isLocalPattern(
  members: CellRef[],
  literalPos: number,
  literalsOnLine: Set<number> | undefined,
  axis: 'row' | 'column',
): boolean {
  const at = (c: CellRef): number => (axis === 'row' ? c.col : c.row);

  let lo = Infinity;
  let hi = -Infinity;
  const occupied = new Set<number>();
  for (const m of members) {
    const p = at(m);
    occupied.add(p);
    if (p < lo) lo = p;
    if (p > hi) hi = p;
  }

  if (hi - lo + 1 > MIN_DENSITY_DIVISOR * members.length) return false;
  // Inside the band, or immediately after it — never immediately before it.
  // The cell in front of a run of formulas is where models legitimately seed a
  // series: an opening balance, a period-zero column, the first month that is
  // actuals before the projection starts. On the real corpus that one idiom
  // accounted for most of what survived the locality rules, and it is a
  // convention rather than a mistake, because derived cells come *after* the
  // inputs they derive from. A total, by contrast, belongs at the end.
  if (literalPos < lo || literalPos > hi + 1) return false;

  for (let p = lo; p <= hi; p++) {
    if (p === literalPos || occupied.has(p)) continue;
    if (literalsOnLine?.has(p)) continue;
    return false;
  }
  return true;
}

function push<K, V>(map: Map<K, V[]>, key: K, value: V): void {
  const list = map.get(key);
  if (list) list.push(value);
  else map.set(key, [value]);
}

function inferFrom(
  engine: Workbook,
  literal: { sheet: number; row: number; col: number; value: Scalar },
  neighbours: CellRef[],
  axis: 'row' | 'column',
  literalsOnLine: Set<number> | undefined,
): HardcodedFinding | undefined {
  if (neighbours.length < 2) return undefined;

  // Group neighbours by the shape of their formula once translated to a common
  // origin: `=SUM(B4:B5)` and `=SUM(C4:C5)` are one shape, not two.
  const shapes = new Map<string, { members: CellRef[]; ast: Node; anchor: CellRef }>();
  for (const n of neighbours) {
    const ast = engine.astAt(literal.sheet, n.row, n.col);
    if (!ast) continue;
    const normalized = normalize(engine, ast, n, literal);
    if (!normalized) continue;
    const key = unparse(normalized);
    const group = shapes.get(key);
    if (group) group.members.push(n);
    else shapes.set(key, { members: [n], ast, anchor: n });
  }

  for (const group of shapes.values()) {
    if (group.members.length < 2) continue;
    // Before translating or evaluating anything: is this shape actually a
    // neighbourhood the literal belongs to? Cheapest check, and the one the
    // detector was missing entirely.
    if (!isLocalPattern(group.members, axis === 'row' ? literal.col : literal.row, literalsOnLine, axis)) {
      continue;
    }

    const dRow = literal.row - group.anchor.row;
    const dCol = literal.col - group.anchor.col;
    const inferred = Workbook.translateShared(group.ast, dRow, dCol);
    if (!sameSeries(engine, literal.sheet, group.ast, inferred)) continue;
    const result = engine.tryEvaluate(literal.sheet, literal.row, literal.col, inferred);
    if (!('value' in result)) continue;
    if (typeof result.value !== 'number' || typeof literal.value !== 'number') continue;

    const stated = literal.value;
    const expected = result.value;
    if (Math.abs(stated - expected) <= EPSILON * Math.max(1, Math.abs(stated), Math.abs(expected))) {
      continue; // the hardcoded number is right — worth nothing to report
    }

    return {
      sheet: literal.sheet,
      sheetName: engine.sheetName(literal.sheet),
      address: `${colLetters(literal.col)}${literal.row}`,
      row: literal.row,
      col: literal.col,
      stated,
      expected,
      formula: unparse(inferred),
      axis,
      pattern: group.members.slice(0, 4).map((m) => `${colLetters(m.col)}${m.row}`),
    };
  }
  return undefined;
}

/** Cells scanned before a range is treated as "too big to characterise". A
 *  whole-column reference tells us nothing about a neighbourhood anyway. */
const MAX_REF_CELLS = 256;

/** What kind of cells a formula reads. Two bits, and they are enough. */
interface RefProfile {
  /** It reads at least one cell that is itself computed. */
  computed: boolean;
  /** It reads at least one cell that exists at all. */
  present: boolean;
}

/**
 * Characterise what a formula would read.
 *
 * Cross-sheet and whole-column references are skipped: an assumption parked on
 * another sheet is shared by every column in a block, so it says nothing about
 * which series a cell belongs to, and a whole-column range says nothing about a
 * neighbourhood.
 */
function referenceProfile(engine: Workbook, sheet: number, node: Node): RefProfile {
  const out: RefProfile = { computed: false, present: false };
  walk(node, (n) => {
    if (n.k !== 'ref' || n.sheet !== undefined) return;
    const a = n.a;
    const b = n.b ?? n.a;
    if (a.colOnly || a.rowOnly || b.colOnly || b.rowOnly) return;
    const r0 = Math.min(a.row, b.row);
    const r1 = Math.max(a.row, b.row);
    const c0 = Math.min(a.col, b.col);
    const c1 = Math.max(a.col, b.col);
    if ((r1 - r0 + 1) * (c1 - c0 + 1) > MAX_REF_CELLS) return;
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        if (engine.record(sheet, r, c) === undefined) continue;
        out.present = true;
        if (engine.astAt(sheet, r, c)) out.computed = true;
      }
    }
  });
  return out;
}

/**
 * Is the literal in the same series as the band, or merely beside it?
 *
 * The locality rules above ask whether the cells are near each other. They
 * cannot tell that `B18:F18` are this year's quarters and `G18` is last year's
 * actuals — a layout so common that the demo model in this repository has one,
 * and the reason two false findings survived the first fix.
 *
 * The tell is what the formulas *read*. Across `B18:F18` the shape is
 * `=x16+x17`, and `B16`/`B17` are themselves computed. Translated onto `G18`
 * the same shape would read `G16`/`G17`, which are hand-entered. A formula that
 * would derive a number from raw inputs, sitting beside formulas that derive
 * theirs from other formulas, is not a missing copy of its neighbours — it is
 * the first cell of a different column of data.
 *
 * The check is symmetric, so it equally protects the ordinary case: `SUM(E4:E5)`
 * over two literals matches `SUM(B4:B5)` over two literals, and the hardcoded
 * total is still reported.
 *
 * The second bit settles a related nonsense the first fix only half-covered: a
 * formula whose inputs *do not exist*. Extrapolating `=B4+B5` upward onto row 3
 * yields `=B1+B2` over two empty title cells, which computes 0 and then
 * "disagrees" with whatever the cell holds. A reconstruction that reads nothing
 * is not a reconstruction; it is an arithmetic accident.
 */
function sameSeries(engine: Workbook, sheet: number, anchor: Node, inferred: Node): boolean {
  const a = referenceProfile(engine, sheet, anchor);
  const b = referenceProfile(engine, sheet, inferred);
  return a.computed === b.computed && a.present === b.present;
}

/** Translate a neighbour's formula onto the literal's position for comparison. */
function normalize(
  engine: Workbook,
  ast: Node,
  from: CellRef,
  to: { row: number; col: number },
): Node | undefined {
  void engine;
  try {
    return Workbook.translateShared(ast, to.row - from.row, to.col - from.col);
  } catch {
    return undefined;
  }
}

function colLetters(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function describeFinding(f: HardcodedFinding): string {
  return (
    `${f.sheetName}!${f.address} states ${plainText(f.stated)}, but the ${f.axis} around it ` +
    `computes ${f.formula} — which is ${plainText(f.expected)}`
  );
}
