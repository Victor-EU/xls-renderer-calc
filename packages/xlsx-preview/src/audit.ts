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
 */

import { unparse, Workbook, type Node, type Scalar } from '@xlscalc/formula-engine';
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

  for (const literal of literals) {
    if (out.length >= limit) break;
    if (typeof literal.value !== 'number') continue;

    const index = bySheet.get(literal.sheet);
    if (!index) continue;

    const finding =
      inferFrom(engine, literal, index.rows.get(literal.row) ?? [], 'row') ??
      inferFrom(engine, literal, index.cols.get(literal.col) ?? [], 'column');
    if (finding) out.push(finding);
  }
  return out;
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

    const dRow = literal.row - group.anchor.row;
    const dCol = literal.col - group.anchor.col;
    const inferred = Workbook.translateShared(group.ast, dRow, dCol);
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
