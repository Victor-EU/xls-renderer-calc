/**
 * Failure, as distinct from error values.
 *
 * `ExcelError` (values.ts) is something Excel itself computes — `#DIV/0!` is a
 * correct answer to `1/0` and renders as `#DIV/0!`. `Unsupported` is *us*
 * admitting we cannot answer. It never becomes a cell value; the evaluator
 * catches it and marks the cell's provenance `unsupported` with the reason
 * verbatim, so the cell renders ⚠ and the gap lands in the coverage backlog.
 *
 * Conflating the two would make the honesty layer meaningless: a user could not
 * tell "your model divides by zero here" from "this renderer gave up here".
 */

export type UnsupportedCode =
  /** Formula text does not parse at all. */
  | 'SYNTAX'
  /** A function name we do not know (may be a real Excel function, or a typo). */
  | 'FN_UNKNOWN'
  /** A function we know of but have not implemented to Excel's semantics. */
  | 'FN_UNIMPLEMENTED'
  /** An argument mode we deliberately refuse (e.g. VLOOKUP approximate match). */
  | 'ARG_UNSUPPORTED'
  /** A reference shape outside what the binder can resolve statically. */
  | 'REF_SHAPE'
  /** A workbook-level feature (1904 dates, external links, array formulas...). */
  | 'FEATURE'
  /** Evaluation exceeded a guard (recursion, iteration, range explosion). */
  | 'LIMIT';

export class Unsupported extends Error {
  override readonly name = 'Unsupported';
  constructor(
    readonly code: UnsupportedCode,
    /** Human-readable, shown in the ⚠ tooltip and aggregated into `gaps`. */
    override readonly message: string,
    /** The function name or ref text, for bucketing the coverage backlog. */
    readonly subject?: string,
  ) {
    super(message);
  }
}

export const unsupported = (code: UnsupportedCode, message: string, subject?: string): never => {
  throw new Unsupported(code, message, subject);
};

export class ParseError extends Error {
  override readonly name = 'ParseError';
  constructor(
    override readonly message: string,
    readonly pos: number,
    readonly formula: string,
  ) {
    super(message);
  }
  /** A one-line caret diagram, for test output and the ⚠ tooltip. */
  annotate(): string {
    return `${this.formula}\n${' '.repeat(Math.max(0, this.pos))}^ ${this.message}`;
  }
}

export const isUnsupported = (e: unknown): e is Unsupported => e instanceof Unsupported;
export const isParseError = (e: unknown): e is ParseError => e instanceof ParseError;
