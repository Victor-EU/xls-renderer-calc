/**
 * The value model.
 *
 * Excel's semantics hinge on distinctions that a naive model collapses, and
 * every collapse produces a wrong number rather than an error. The three that
 * matter most, encoded here as distinct types:
 *
 *   1. **Blank is not empty string and not zero.** An empty cell in arithmetic
 *      is 0, in comparison it equals *both* 0 and "", and in COUNT it is
 *      invisible. `""` does none of that. Blank is `null`; `""` is a string.
 *   2. **An Excel error is a value, not a failure.** `#DIV/0!` propagates
 *      through operators, is trapped by IFERROR, and renders as itself with
 *      provenance `computed`. Only *our* inability to evaluate is a failure,
 *      and that is signalled by throwing `Unsupported`, never by an error value.
 *   3. **A reference is not its value.** `SUM(A1:A3)` ignores text inside the
 *      range; `SUM("5")` coerces it. Same function, different argument kind, so
 *      references survive as their own type until a function decides.
 *
 * Dates are IEEE-754 serials, never `Date` objects — conversion happens at the
 * render boundary only (see `dates.ts` for the 1900 leap-year bug).
 */

export type ErrorKind =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#SPILL!'
  | '#CALC!';

/** A computed Excel error value. Distinct from `Unsupported` (see errors.ts). */
export class ExcelError {
  constructor(
    readonly kind: ErrorKind,
    /** Optional human detail — surfaced in tooltips, never in the cell text. */
    readonly detail?: string,
  ) {}
  toString(): string {
    return this.kind;
  }
}

export const ERR = {
  null: new ExcelError('#NULL!'),
  div0: new ExcelError('#DIV/0!'),
  value: new ExcelError('#VALUE!'),
  ref: new ExcelError('#REF!'),
  name: new ExcelError('#NAME?'),
  num: new ExcelError('#NUM!'),
  na: new ExcelError('#N/A'),
  spill: new ExcelError('#SPILL!'),
  calc: new ExcelError('#CALC!'),
} as const;

const ERR_BY_KIND: Record<ErrorKind, ExcelError> = {
  '#NULL!': ERR.null,
  '#DIV/0!': ERR.div0,
  '#VALUE!': ERR.value,
  '#REF!': ERR.ref,
  '#NAME?': ERR.name,
  '#NUM!': ERR.num,
  '#N/A': ERR.na,
  '#SPILL!': ERR.spill,
  '#CALC!': ERR.calc,
};

export function errorOf(kind: string): ExcelError | undefined {
  return ERR_BY_KIND[kind as ErrorKind];
}

/** `null` is BLANK — an empty cell. It is not `""` and it is not `0`. */
export type Blank = null;

/** What a cell can hold and what most functions consume. */
export type Scalar = number | string | boolean | ExcelError | Blank;

/**
 * A rectangular block of scalars: an array constant (`{1,2;3,4}`), a
 * materialised range, or an array-valued function result. Row-major, 0-based
 * internally; every public row/col elsewhere in the engine is 1-based.
 */
export class Matrix {
  constructor(
    readonly rows: number,
    readonly cols: number,
    readonly data: Scalar[],
  ) {}

  static of(values: Scalar[][]): Matrix {
    const rows = values.length;
    const cols = rows === 0 ? 0 : Math.max(...values.map((r) => r.length));
    const data: Scalar[] = new Array(rows * cols).fill(null);
    for (let r = 0; r < rows; r++) {
      const row = values[r]!;
      for (let c = 0; c < cols; c++) data[r * cols + c] = c < row.length ? row[c]! : null;
    }
    return new Matrix(rows, cols, data);
  }

  static scalar(v: Scalar): Matrix {
    return new Matrix(1, 1, [v]);
  }

  /** 0-based access. Out of bounds yields `#N/A`, matching array broadcasting. */
  at(r: number, c: number): Scalar {
    if (r < 0 || c < 0 || r >= this.rows || c >= this.cols) return ERR.na;
    return this.data[r * this.cols + c]!;
  }

  get size(): number {
    return this.rows * this.cols;
  }

  *values(): IterableIterator<Scalar> {
    for (const v of this.data) yield v;
  }
}

/**
 * A live reference into the workbook — a rectangle on one sheet. Kept unresolved
 * so functions can see "this argument was a range" (SUM skips text) versus
 * "this argument was a value" (SUM coerces numeric text).
 *
 * All coordinates 1-based and inclusive. `sheet` is a sheet index, not a name;
 * name resolution happens during parse-binding.
 */
export class RefValue {
  constructor(
    readonly sheet: number,
    readonly top: number,
    readonly left: number,
    readonly bottom: number,
    readonly right: number,
    /**
     * The extent the formula actually named, when it is larger than the one we
     * will materialise.
     *
     * `A:A` names 1,048,576 rows. We clamp `bottom` to the used range, because
     * reading the rest is unaffordable and iterating it is pointless — the
     * cells are empty, so `SUM`, `COUNT` and `MATCH` get the same answer either
     * way. That clamp is an implementation detail of *iteration*, and two kinds
     * of function must not see it:
     *
     *   - those that report an extent — Excel's `ROWS(A:A)` is 1048576, not
     *     however many rows happen to be occupied;
     *   - those that index by position — `INDEX(Sheet!D:O, 63, 1)` is the cell
     *     D63 whether or not anything has ever been written near it.
     *
     * Conflating the two made a board pack render `#REF!` where Excel finds an
     * empty cell. Keeping the declared extent beside the clamped one lets each
     * function ask for the one it actually means.
     */
    readonly declared?: { bottom: number; right: number },
  ) {}

  /** The last row the formula named, which may be far below the used range. */
  get declaredBottom(): number {
    return this.declared?.bottom ?? this.bottom;
  }
  get declaredRight(): number {
    return this.declared?.right ?? this.right;
  }
  get declaredRows(): number {
    return this.declaredBottom - this.top + 1;
  }
  get declaredCols(): number {
    return this.declaredRight - this.left + 1;
  }
  /** True when the clamp actually discarded something. */
  get isClamped(): boolean {
    return this.declaredBottom !== this.bottom || this.declaredRight !== this.right;
  }

  static cell(sheet: number, row: number, col: number): RefValue {
    return new RefValue(sheet, row, col, row, col);
  }

  get rows(): number {
    return this.bottom - this.top + 1;
  }
  get cols(): number {
    return this.right - this.left + 1;
  }
  get isSingleCell(): boolean {
    return this.rows === 1 && this.cols === 1;
  }
  get cellCount(): number {
    return this.rows * this.cols;
  }
}

/** Anything an expression node can evaluate to. */
export type EvalValue = Scalar | Matrix | RefValue;

export const isErr = (v: unknown): v is ExcelError => v instanceof ExcelError;
export const isBlank = (v: unknown): v is Blank => v === null;
export const isMatrix = (v: unknown): v is Matrix => v instanceof Matrix;
export const isRef = (v: unknown): v is RefValue => v instanceof RefValue;
export const isScalar = (v: EvalValue): v is Scalar => !isMatrix(v) && !isRef(v);

/**
 * Excel's cross-type ordering for comparison and sorting: number < text <
 * FALSE < TRUE. Blank is not ranked here — it is context-coerced first.
 */
export function typeRank(v: Scalar): number {
  if (typeof v === 'number') return 1;
  if (typeof v === 'string') return 2;
  if (typeof v === 'boolean') return 3;
  return 4; // errors never reach ordering; they short-circuit
}
