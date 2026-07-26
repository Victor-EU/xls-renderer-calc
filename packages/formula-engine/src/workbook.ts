/**
 * The workbook: storage, dependency graph, evaluation order, provenance.
 *
 * Three decisions here carry most of the weight.
 *
 * **Computed values never overwrite the file's cached ones.** Both are kept, so
 * every rendered number knows where it came from, and the disagreement between
 * them is available as an audit finding rather than being destroyed on load.
 *
 * **Unsupported poisons downstream.** If a cell cannot be computed, every cell
 * that reads it is also uncomputable and is marked so. Letting a `SUM` quietly
 * treat an uncomputable precedent as zero would produce exactly the confident
 * wrong subtotal this engine exists to prevent.
 *
 * **Cycles are detected, never iterated.** Excel refuses circular references by
 * default and only iterates when a user opts in; a renderer that silently
 * iterated would show numbers no one asked for.
 */

import { MAX_COLS, cellId, formatAddress, unpackId } from './a1.js';
import { translate, walk, type Node } from './ast.js';
import { isParseError, isUnsupported, Unsupported } from './errors.js';
import { installStandardLibrary } from './functions/index.js';
import { evaluate, scalarOf, type EvalHost, type FnContext } from './interpreter.js';
import { parseFormula } from './parser.js';
import { ERR, RefValue, isMatrix, isRef, type EvalValue, type Scalar } from './values.js';

installStandardLibrary();

export type Provenance =
  /** Straight from the file's `<v>` — we did not touch it. */
  | 'literal'
  /** Read from the file's formula cache; we did not recompute it. */
  | 'cached'
  /** We evaluated it. */
  | 'computed'
  /** Outside engine capability — rendered ⚠, never guessed. */
  | 'unsupported'
  /** Part of, or downstream of, a reference cycle. */
  | 'circular'
  /** Computed, but depends on NOW/TODAY/RAND, so it is frozen at load time. */
  | 'volatile';

export interface FormulaCell {
  readonly src: string;
  ast?: Node;
  /** Why the formula could not be used at all — parse failure or refusal. */
  fault?: { code: string; message: string; subject?: string };
}

export interface CellRecord {
  value: Scalar;
  provenance: Provenance;
  /** The `<v>` the file carried, retained for the computed-vs-cached diff. */
  cached?: Scalar;
  reason?: string;
}

export class Sheet {
  readonly cells = new Map<number, Scalar>();
  readonly formulas = new Map<number, FormulaCell>();
  maxRow = 0;
  maxCol = 0;

  constructor(
    readonly name: string,
    readonly index: number,
  ) {}

  key(row: number, col: number): number {
    return row * MAX_COLS + col;
  }

  note(row: number, col: number): void {
    if (row > this.maxRow) this.maxRow = row;
    if (col > this.maxCol) this.maxCol = col;
  }
}

export interface EvalStats {
  formulas: number;
  computed: number;
  cached: number;
  unsupported: number;
  circular: number;
  volatile: number;
  /** Cells where our computed value disagrees with the file's cache. */
  mismatched: number;
  totalMs: number;
}

export interface Gap {
  code: string;
  subject: string;
  reason: string;
  count: number;
  sample: string;
}

export interface EvalReport {
  stats: EvalStats;
  gaps: Gap[];
  /** Cells whose computed value differs from the cached one, for the audit view. */
  mismatches: Array<{ address: string; computed: Scalar; cached: Scalar }>;
}

export interface WorkbookOptions {
  date1904?: boolean;
  /**
   * Clock for NOW/TODAY, read exactly once per run so evaluation is
   * reproducible — never `Date.now()` inside a function body.
   */
  now?: number;
}

export class Workbook implements EvalHost {
  readonly sheets: Sheet[] = [];
  private readonly byName = new Map<string, Sheet>();
  private readonly names = new Map<string, { sheet?: number; node: Node }>();
  private readonly records = new Map<number, CellRecord>();
  readonly date1904: boolean;
  private nowSerial: number;

  /** Set while evaluating so a read of an uncomputable cell can unwind the reader. */
  private evaluating = false;

  constructor(opts: WorkbookOptions = {}) {
    this.date1904 = opts.date1904 ?? false;
    this.nowSerial = opts.now ?? serialNow(this.date1904);
  }

  // ---------------------------------------------------------------- authoring

  addSheet(name: string): Sheet {
    const sheet = new Sheet(name, this.sheets.length);
    this.sheets.push(sheet);
    this.byName.set(name.toLowerCase(), sheet);
    return sheet;
  }

  sheet(nameOrIndex: string | number): Sheet | undefined {
    return typeof nameOrIndex === 'number'
      ? this.sheets[nameOrIndex]
      : this.byName.get(nameOrIndex.toLowerCase());
  }

  setValue(sheet: number, row: number, col: number, value: Scalar): void {
    const s = this.sheets[sheet];
    if (!s) throw new Error(`no sheet ${sheet}`);
    s.cells.set(s.key(row, col), value);
    s.note(row, col);
    this.records.set(cellId(sheet, row, col), { value, provenance: 'literal' });
  }

  /**
   * Attach a formula. `cached` is the value the file carried, kept for the diff
   * and used as a fallback when the formula turns out to be uncomputable.
   */
  setFormula(sheet: number, row: number, col: number, src: string, cached?: Scalar): void {
    const s = this.sheets[sheet];
    if (!s) throw new Error(`no sheet ${sheet}`);
    const key = s.key(row, col);
    s.formulas.set(key, { src });
    s.cells.set(key, cached ?? null);
    s.note(row, col);
    const record: CellRecord =
      cached === undefined
        ? { value: null, provenance: 'unsupported', reason: 'not evaluated yet' }
        : { value: cached, provenance: 'cached', cached };
    this.records.set(cellId(sheet, row, col), record);
  }

  defineName(name: string, formula: string, sheet?: number): void {
    try {
      const node = parseFormula(formula);
      this.names.set(nameKey(name, sheet), sheet === undefined ? { node } : { sheet, node });
    } catch {
      // A name we cannot parse simply stays undefined; formulas using it will
      // report the failure themselves rather than silently resolving to zero.
    }
  }

  // ------------------------------------------------------------------ reading

  record(sheet: number, row: number, col: number): CellRecord | undefined {
    return this.records.get(cellId(sheet, row, col));
  }

  formulaAt(sheet: number, row: number, col: number): string | undefined {
    const s = this.sheets[sheet];
    return s?.formulas.get(s.key(row, col))?.src;
  }

  /**
   * The cells a formula reads directly — "explain this number", one hop.
   *
   * With a dependency graph this is the same capability as Excel's Trace
   * Precedents, and it is the natural pairing with an agent that *wrote* the
   * model: the stated rationale and the graph's actual structure can be compared.
   * Only available after `evaluateAll()` has parsed the formulas.
   */
  precedents(sheet: number, row: number, col: number, limit = 500): Array<{ sheet: number; row: number; col: number }> {
    const s = this.sheets[sheet];
    const ast = s?.formulas.get(s.key(row, col))?.ast;
    if (!ast) return [];
    const out: Array<{ sheet: number; row: number; col: number }> = [];
    for (const rect of this.precedentRects(ast, sheet, row, col)) {
      for (let r = rect.top; r <= rect.bottom && out.length < limit; r++) {
        for (let c = rect.left; c <= rect.right && out.length < limit; c++) {
          out.push({ sheet: rect.sheet, row: r, col: c });
        }
      }
    }
    return out;
  }

  // -------------------------------------------------------------- EvalHost

  sheetIndex(name: string): number | undefined {
    return this.byName.get(name.toLowerCase())?.index;
  }

  sheetName(index: number): string {
    return this.sheets[index]?.name ?? `#${index}`;
  }

  usedRange(sheet: number): { rows: number; cols: number } {
    const s = this.sheets[sheet];
    return { rows: s?.maxRow ?? 0, cols: s?.maxCol ?? 0 };
  }

  now(): number {
    return this.nowSerial;
  }

  cellValue(sheet: number, row: number, col: number): Scalar {
    const rec = this.records.get(cellId(sheet, row, col));
    if (!rec) return null;
    if (this.evaluating && (rec.provenance === 'unsupported' || rec.provenance === 'circular')) {
      throw new Unsupported(
        rec.provenance === 'circular' ? 'LIMIT' : 'FEATURE',
        `depends on ${formatAddress(this.sheetName(sheet), row, col)}, which could not be computed`,
        'dependency',
      );
    }
    return rec.value;
  }

  private readonly resolving = new Set<string>();

  definedName(name: string, sheet: number): RefValue | Scalar | undefined {
    const key = nameKey(name, sheet);
    const entry = this.names.get(key) ?? this.names.get(nameKey(name, undefined));
    if (!entry) return undefined;
    // A name defined in terms of itself would recurse forever; that is a cycle,
    // and cycles are reported, never followed.
    if (this.resolving.has(key)) return ERR.ref;
    this.resolving.add(key);
    try {
      const ctx = this.contextFor(entry.sheet ?? sheet, 1, 1);
      const v = evaluate(entry.node, ctx);
      if (isRef(v)) return v;
      if (isMatrix(v)) return v.size ? v.at(0, 0) : ERR.value;
      return v as Scalar;
    } finally {
      this.resolving.delete(key);
    }
  }

  private contextFor(sheet: number, row: number, col: number): FnContext {
    return { host: this, sheet, row, col, arrayMode: false, volatile: false };
  }

  // --------------------------------------------------------------- evaluation

  /**
   * Compute every formula in the workbook, in dependency order, and return what
   * happened. Safe to call repeatedly; each run starts from the file's values.
   */
  evaluateAll(): EvalReport {
    const started = Date.now();
    const stats: EvalStats = {
      formulas: 0,
      computed: 0,
      cached: 0,
      unsupported: 0,
      circular: 0,
      volatile: 0,
      mismatched: 0,
      totalMs: 0,
    };
    const gaps = new Map<string, Gap>();
    const mismatches: EvalReport['mismatches'] = [];

    const noteGap = (code: string, subject: string, reason: string, id: number): void => {
      const k = `${code}:${subject}`;
      const existing = gaps.get(k);
      if (existing) existing.count++;
      else gaps.set(k, { code, subject, reason, count: 1, sample: this.addressOf(id) });
    };

    // 1. Parse. A formula that will not parse is a per-cell refusal, not a crash.
    const nodes = new Map<number, Node>();
    for (const sheet of this.sheets) {
      for (const [key, cell] of sheet.formulas) {
        stats.formulas++;
        const row = Math.floor(key / MAX_COLS);
        const col = key - row * MAX_COLS;
        const id = cellId(sheet.index, row, col);
        try {
          const ast = parseFormula(cell.src);
          cell.ast = ast;
          nodes.set(id, ast);
        } catch (e) {
          const { code, subject, message } = classify(e);
          cell.fault = { code, message, subject };
          this.markUnsupported(id, message);
          noteGap(code, subject, message, id);
        }
      }
    }

    // 2. Evaluation order, without materialising the dependency graph.
    //
    //    Storing an edge per precedent cell looks natural and is a trap: a model
    //    with a running total (`SUM($B$1:B7)` on every row) has O(rows²) edges,
    //    and at 5,000 rows that is 12.5 million Set entries — hundreds of
    //    megabytes, which took a 45,000-formula workbook from "slow" to
    //    "out of memory".
    //
    //    Instead the precedent *rectangles* are kept per formula and walked
    //    lazily by an iterative depth-first search, whose post-order is a
    //    topological order. Memory is O(formulas); the walk allocates nothing
    //    per edge. The search is iterative rather than recursive because a
    //    row-to-row chain is as deep as the model is long, and that overflows
    //    the call stack long before it exhausts an array.
    const rectsOf = new Map<number, RefValue[]>();
    const selfRefs = new Set<number>();
    for (const [id, ast] of nodes) {
      const { sheet, row, col } = unpackId(id);
      const rects = this.precedentRects(ast, sheet, row, col);
      rectsOf.set(id, rects);
      for (const rect of rects) {
        if (
          rect.sheet === sheet &&
          row >= rect.top &&
          row <= rect.bottom &&
          col >= rect.left &&
          col <= rect.right
        ) {
          selfRefs.add(id);
        }
      }
    }

    const WHITE = 0;
    const GREY = 1;
    const BLACK = 2;
    const state = new Map<number, number>();
    const onCycle = new Set<number>(selfRefs);
    const order: number[] = [];

    interface Frame {
      id: number;
      rects: RefValue[];
      ri: number;
      r: number;
      c: number;
    }
    const frameFor = (id: number): Frame => {
      const rects = rectsOf.get(id) ?? [];
      const first = rects[0];
      return { id, rects, ri: 0, r: first?.top ?? 0, c: first?.left ?? 0 };
    };

    /** Next precedent of `f` that is itself a formula, advancing the cursor. */
    const nextPrecedent = (f: Frame): number | undefined => {
      while (f.ri < f.rects.length) {
        const rect = f.rects[f.ri]!;
        while (f.r <= rect.bottom) {
          while (f.c <= rect.right) {
            const pid = cellId(rect.sheet, f.r, f.c);
            f.c++;
            if (pid !== f.id && nodes.has(pid)) return pid;
          }
          f.c = rect.left;
          f.r++;
        }
        f.ri++;
        const nextRect = f.rects[f.ri];
        f.r = nextRect?.top ?? 0;
        f.c = nextRect?.left ?? 0;
      }
      return undefined;
    };

    for (const start of nodes.keys()) {
      if ((state.get(start) ?? WHITE) !== WHITE) continue;
      const stack: Frame[] = [frameFor(start)];
      state.set(start, GREY);

      while (stack.length > 0) {
        const frame = stack[stack.length - 1]!;
        const next = nextPrecedent(frame);
        if (next === undefined) {
          state.set(frame.id, BLACK);
          order.push(frame.id);
          stack.pop();
          continue;
        }
        const seen = state.get(next) ?? WHITE;
        if (seen === BLACK) continue;
        if (seen === GREY) {
          // Everything from `next` up to the top of the stack is in the cycle.
          for (let i = stack.length - 1; i >= 0; i--) {
            onCycle.add(stack[i]!.id);
            if (stack[i]!.id === next) break;
          }
          continue;
        }
        state.set(next, GREY);
        stack.push(frameFor(next));
      }
    }

    // 3. Cells on a cycle are reported, never iterated to some fixed point.
    for (const id of onCycle) {
      this.markCircular(id);
      stats.circular++;
      noteGap('CYCLE', 'circular reference', 'this cell is part of a reference cycle', id);
    }

    // 4. Evaluate in order.
    this.evaluating = true;
    try {
      for (const id of order) {
        if (onCycle.has(id)) continue;
        const ast = nodes.get(id)!;
        const { sheet, row, col } = unpackId(id);
        const ctx = this.contextFor(sheet, row, col);
        const rec = this.records.get(id)!;
        try {
          const raw = evaluate(ast, ctx);
          const value = collapse(raw, ctx);
          rec.value = value;
          rec.provenance = ctx.volatile ? 'volatile' : 'computed';
          delete rec.reason;
          if (ctx.volatile) stats.volatile++;
          else stats.computed++;

          if (rec.cached !== undefined && !valuesAgree(value, rec.cached) && !ctx.volatile) {
            stats.mismatched++;
            mismatches.push({ address: this.addressOf(id), computed: value, cached: rec.cached });
          }
        } catch (e) {
          const { code, subject, message } = classify(e);
          this.markUnsupported(id, message);
          noteGap(code, subject, message, id);
        }
      }
    } finally {
      this.evaluating = false;
    }

    for (const rec of this.records.values()) {
      if (rec.provenance === 'unsupported') stats.unsupported++;
      else if (rec.provenance === 'cached') stats.cached++;
    }
    stats.totalMs = Date.now() - started;

    return { stats, gaps: [...gaps.values()].sort((a, b) => b.count - a.count), mismatches };
  }

  // ------------------------------------------------------------------ helpers

  private addressOf(id: number): string {
    const { sheet, row, col } = unpackId(id);
    return formatAddress(this.sheetName(sheet), row, col);
  }

  private markUnsupported(id: number, reason: string): void {
    const rec = this.records.get(id) ?? { value: null, provenance: 'unsupported' as Provenance };
    rec.provenance = 'unsupported';
    rec.reason = reason;
    // Keep the file's cached value visible in the record so the UI can offer it,
    // but the *rendered* value for an unsupported cell is ⚠, decided by the view.
    rec.value = rec.cached ?? null;
    this.records.set(id, rec);
  }

  private markCircular(id: number): void {
    const rec = this.records.get(id) ?? { value: null, provenance: 'circular' as Provenance };
    rec.provenance = 'circular';
    rec.reason = 'circular reference';
    this.records.set(id, rec);
  }

  /** Every rectangle a formula reads, with whole-column refs already clamped. */
  private precedentRects(ast: Node, sheet: number, row: number, col: number): RefValue[] {
    const out: RefValue[] = [];
    const ctx = this.contextFor(sheet, row, col);
    walk(ast, (n) => {
      if (n.k === 'ref') {
        try {
          const v = evaluate(n, ctx);
          if (isRef(v)) out.push(v);
        } catch {
          // A reference we cannot resolve is reported when the cell evaluates.
        }
      } else if (n.k === 'name') {
        const resolved = this.definedName(n.name, n.sheet ? (this.sheetIndex(n.sheet) ?? sheet) : sheet);
        if (isRef(resolved)) out.push(resolved);
      }
    });
    return out;
  }

  /** Apply a shared formula to a sibling cell by translating its relative refs. */
  static translateShared(master: Node, dRow: number, dCol: number): Node {
    return translate(master, dRow, dCol, 1_048_576, MAX_COLS);
  }

  /** The parsed formula of a cell, available after `evaluateAll()`. */
  astAt(sheet: number, row: number, col: number): Node | undefined {
    const s = this.sheets[sheet];
    return s?.formulas.get(s.key(row, col))?.ast;
  }

  /**
   * Evaluate a formula *as if* it lived at a cell, without storing it.
   *
   * Used by the audit pass to ask "what would this cell hold if it summed its
   * neighbours, like the rest of the row does?" — and it is the same primitive a
   * what-if edit needs later.
   */
  tryEvaluate(
    sheet: number,
    row: number,
    col: number,
    formula: string | Node,
  ): { value: Scalar } | { unsupported: string } {
    try {
      const ast = typeof formula === 'string' ? parseFormula(formula) : formula;
      const ctx = this.contextFor(sheet, row, col);
      return { value: collapse(evaluate(ast, ctx), ctx) };
    } catch (e) {
      // Same rule as `evaluateAll`: this is a speculative evaluation on behalf
      // of the audit pass, and it must never be able to fail the load.
      return { unsupported: classify(e).message };
    }
  }
}

/**
 * Turn anything thrown while handling one cell into that cell's refusal.
 *
 * `Unsupported` and `ParseError` are the two we author, and they carry their own
 * code and subject. Everything else is a defect in this engine — and the rule
 * that matters is that it stays *this cell's* problem. Rethrowing unwound
 * `evaluateAll` and lost the entire workbook, which was reachable from a single
 * ordinary formula: `=VLOOKUP(A1,T,2,)` threw a TypeError and `=MIN(A:A)` down a
 * long column threw a RangeError, and either one turned a 138,000-formula
 * preview into a blank screen. Failing one cell loudly is strictly better than
 * failing all of them silently, so unexpected failures are filed under
 * `INTERNAL` — visible in `gaps`, impossible to mistake for a decision.
 */
function classify(e: unknown): { code: string; subject: string; message: string } {
  if (isUnsupported(e)) return { code: e.code, subject: e.subject ?? e.code, message: e.message };
  if (isParseError(e)) return { code: 'SYNTAX', subject: 'syntax', message: e.message };
  const detail = e instanceof Error ? `${e.name}: ${e.message}` : String(e);
  return {
    code: 'INTERNAL',
    subject: e instanceof Error ? e.name : 'internal',
    message: `internal error while evaluating this cell — ${detail}`,
  };
}

function nameKey(name: string, sheet: number | undefined): string {
  return `${sheet ?? '*'}:${name.toLowerCase()}`;
}

/**
 * A formula entered in a single cell keeps one value: the top-left of an array
 * result, or the implicit intersection of a bare range — the legacy behaviour,
 * which is the right one here because this engine renders saved files and a
 * saved file already contains whatever the writer spilled.
 */
function collapse(v: EvalValue, ctx: FnContext): Scalar {
  const out = isMatrix(v)
    ? v.size === 0
      ? ERR.value
      : v.at(0, 0)
    : isRef(v)
      ? scalarOf(v, ctx)
      : (v as Scalar);
  // Blank is a value *inside* an expression — that is what makes `IF(G9=0,…)`
  // work over an empty column. But a cell never holds blank as a result: Excel
  // renders `=J1` over an empty J1 as 0, not as an empty cell.
  return out === null ? 0 : out;
}

/**
 * Compare a computed value with the file's cached one.
 *
 * Floats get a *relative* epsilon: LibreOffice writes 14 significant digits into
 * `<v>` where the double carries 17, so `===` would report a mismatch on every
 * cell in the file.
 */
export function valuesAgree(a: Scalar, b: Scalar): boolean {
  if (typeof a === 'number' && typeof b === 'number') {
    return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  }
  if (a === null || b === null) return a === b;
  if (typeof a === 'object' || typeof b === 'object') return String(a) === String(b);
  return a === b;
}

function serialNow(date1904: boolean): number {
  const ms = Date.now();
  const epoch = date1904 ? 24107 : 25569;
  return ms / 86_400_000 + epoch;
}
