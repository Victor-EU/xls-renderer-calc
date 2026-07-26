/**
 * Pratt parser for Excel formulas.
 *
 * Two precedence facts here are not folklore — they are the reason a
 * hand-written evaluator quietly disagrees with Excel:
 *
 *   `-2^2  = 4`   unary minus binds *tighter* than `^` (Microsoft's documented
 *                 order is: reference ops, negation, %, ^, * /, + -, &, compare),
 *                 unlike mathematics and unlike every C-family language.
 *   `2^3^2 = 64`  `^` is **left**-associative in Excel, not right.
 *
 * Both are asserted in parser.test.ts and re-asserted against LibreOffice in the
 * oracle harness, because getting them wrong produces a plausible number rather
 * than an error — the worst failure mode this project has.
 *
 * Anything the parser cannot represent faithfully is thrown as `Unsupported`
 * with a reason, never coerced into something that happens to parse.
 */

import {
  MAX_COLS,
  MAX_ROWS,
  parseCellRef,
  parseColOnly,
  parseRowOnly,
  type CellRef,
} from './a1.js';
import {
  BP_ISECT,
  BP_PCT,
  BP_UNARY,
  INFIX_BP,
  type BinOp,
  type Node,
} from './ast.js';
import { ParseError, Unsupported } from './errors.js';
import { tokenize, type Token } from './lexer.js';
import { ERR, errorOf } from './values.js';

// The binding powers live in ast.ts, because `unparse` has to be this parser's
// exact inverse and a second copy of the table is a second chance to be wrong.

/** Function-name namespaces Excel writes into the file but never displays. */
const FN_PREFIXES = ['_XLFN.', '_XLWS.'];

const OPERAND_START = new Set<Token['type']>(['num', 'str', 'err', 'ident', 'sheet', 'lparen', 'lbrace', 'at']);

class Parser {
  private p = 0;
  constructor(
    private readonly toks: Token[],
    private readonly src: string,
  ) {}

  private cur(): Token {
    return this.toks[this.p]!;
  }
  private skipWs(): void {
    while (this.cur().type === 'ws') this.p++;
  }
  /**
   * Non-destructive: it must NOT consume the whitespace it looks past, because
   * whitespace between two operands is the intersection operator and parseExpr
   * needs to still see it.
   */
  private peek(): Token {
    return this.lookahead(0);
  }
  private take(): Token {
    this.skipWs();
    return this.toks[this.p++]!;
  }
  /** Consume the token `peek()` just reported. */
  private bump(): void {
    this.skipWs();
    this.p++;
  }
  /** Look ahead past whitespace by `n` significant tokens without consuming. */
  private lookahead(n: number): Token {
    let i = this.p;
    let seen = 0;
    while (i < this.toks.length) {
      const t = this.toks[i]!;
      if (t.type !== 'ws') {
        if (seen === n) return t;
        seen++;
      }
      i++;
    }
    return this.toks[this.toks.length - 1]!;
  }
  private fail(msg: string, pos?: number): never {
    throw new ParseError(msg, pos ?? this.cur().pos, this.src);
  }
  private expect(type: Token['type'], what: string): Token {
    const t = this.take();
    if (t.type !== type) this.fail(`expected ${what}`, t.pos);
    return t;
  }

  parse(): Node {
    const node = this.parseExpr(0);
    const t = this.peek();
    if (t.type !== 'eof') this.fail(`unexpected ${t.text || 'end of input'}`, t.pos);
    return node;
  }

  parseExpr(minBp: number): Node {
    let left = this.parseUnary();

    for (;;) {
      const sawWs = this.cur().type === 'ws';
      this.skipWs();
      const t = this.cur();

      // Excel's intersection operator is a space between two operands.
      if (sawWs && OPERAND_START.has(t.type) && BP_ISECT >= minBp) {
        const right = this.parseExpr(BP_ISECT + 1);
        left = { k: 'isect', l: left, r: right };
        continue;
      }
      if (t.type !== 'op') break;

      if (t.text === '%') {
        if (BP_PCT < minBp) break;
        this.bump();
        left = { k: 'pct', v: left };
        continue;
      }

      const bp = INFIX_BP[t.text as BinOp];
      if (bp === undefined || bp < minBp) break;
      this.bump();
      // Every value operator in Excel is left-associative, `^` included, so the
      // right operand is parsed at bp + 1.
      const right = this.parseExpr(bp + 1);
      left = { k: 'bin', op: t.text as BinOp, l: left, r: right };
    }
    return left;
  }

  private parseUnary(): Node {
    const t = this.peek();
    if (t.type === 'op' && (t.text === '-' || t.text === '+')) {
      this.bump();
      const v = this.parseExpr(BP_UNARY);
      return { k: 'un', op: t.text, v };
    }
    return this.parsePrimary();
  }

  private parsePrimary(): Node {
    const t = this.take();
    switch (t.type) {
      case 'num': {
        // `1:1` is a whole-row reference, not the number 1 followed by junk.
        if (this.peek().type === 'op' && this.peek().text === ':') {
          return this.finishRowRange(t, undefined);
        }
        return { k: 'num', v: Number(t.text) };
      }
      case 'str':
        return { k: 'str', v: t.text };
      case 'err': {
        const e = errorOf(t.text);
        if (!e) throw new Unsupported('FEATURE', `error literal ${t.text} is not supported`, t.text);
        return { k: 'err', v: e };
      }
      case 'at':
        return { k: 'at', v: this.parsePrimary() };
      case 'lparen': {
        const parts: Node[] = [this.parseExpr(0)];
        while (this.peek().type === 'comma') {
          this.bump();
          parts.push(this.parseExpr(0));
        }
        this.expect('rparen', ')');
        return parts.length === 1 ? parts[0]! : { k: 'union', parts };
      }
      case 'lbrace':
        return this.parseArrayConstant(t);
      case 'bracket':
        throw new Unsupported(
          'REF_SHAPE',
          `external workbook links are not supported (${t.text})`,
          t.text,
        );
      case 'sheet': {
        this.expect('bang', "'!' after a sheet name");
        return this.parseRefBody(t.text);
      }
      case 'ident':
        return this.parseIdent(t);
      default:
        this.fail(`unexpected ${t.text || 'end of formula'}`, t.pos);
    }
  }

  private parseIdent(t: Token): Node {
    const nxt = this.peek();

    // Function call: NAME(
    if (nxt.type === 'lparen') {
      this.bump();
      const args = this.parseArgs();
      return { k: 'fn', name: normalizeFnName(t.text), args, pos: t.pos };
    }

    // Structured table reference: Table1[Revenue]
    if (nxt.type === 'bracket') {
      throw new Unsupported(
        'REF_SHAPE',
        `structured table references are not supported (${t.text}${nxt.text})`,
        `${t.text}${nxt.text}`,
      );
    }

    // Sheet-qualified: Sheet1!A1 or Sheet1!MyName
    if (nxt.type === 'bang') {
      this.bump();
      return this.parseRefBody(t.text);
    }

    // 3D reference: Sheet1:Sheet3!A1 — distinguishable from A1:B2 by the '!'
    if (nxt.type === 'op' && nxt.text === ':' && this.lookahead(2).type === 'bang') {
      throw new Unsupported('REF_SHAPE', `3D references across sheets are not supported (${t.text}:…)`);
    }

    return this.identToRefOrName(t, undefined);
  }

  /** Parse the address part after an optional `Sheet!` prefix. */
  private parseRefBody(sheet: string): Node {
    const t = this.take();
    if (t.type === 'err') {
      const e = errorOf(t.text);
      return { k: 'err', v: e ?? ERR.ref };
    }
    if (t.type === 'num') return this.finishRowRange(t, sheet);
    if (t.type !== 'ident') this.fail('expected a cell reference or name after !', t.pos);
    return this.identToRefOrName(t, sheet);
  }

  private identToRefOrName(t: Token, sheet: string | undefined): Node {
    const nxt = this.peek();
    const isRangeOp = nxt.type === 'op' && nxt.text === ':';

    if (isRangeOp) {
      this.bump();
      let secondSheet: string | undefined;
      let second = this.take();
      // Sheet1!A1:Sheet1!B2 — legal, but only when both sides name one sheet.
      if ((second.type === 'ident' || second.type === 'sheet') && this.peek().type === 'bang') {
        secondSheet = second.text;
        this.bump();
        second = this.take();
      }
      if (second.type === 'err') return { k: 'err', v: errorOf(second.text) ?? ERR.ref };
      if (second.type !== 'ident' && second.type !== 'num') {
        this.fail('expected a cell reference after :', second.pos);
      }
      if (secondSheet !== undefined && secondSheet !== sheet) {
        throw new Unsupported(
          'REF_SHAPE',
          `a range may not span sheets (${sheet ?? ''}!…:${secondSheet}!…)`,
        );
      }
      const a = cornerOf(t.text);
      const b = cornerOf(second.text);
      if (!a || !b) this.fail(`not a valid reference: ${t.text}:${second.text}`, t.pos);
      return this.makeRange(a, b, sheet, t);
    }

    const cell = parseCellRef(t.text);
    if (cell) return { k: 'ref', sheet, a: cell };

    if (/^(TRUE|FALSE)$/i.test(t.text) && sheet === undefined) {
      return { k: 'bool', v: t.text.toUpperCase() === 'TRUE' };
    }
    // Not an address — a defined name. Resolution (and #NAME? for one that does
    // not exist in the workbook) happens at bind time, where we can actually
    // check, rather than being guessed here.
    return { k: 'name', name: t.text.replace(/^\$/, ''), sheet };
  }

  private finishRowRange(first: Token, sheet: string | undefined): Node {
    // consume ':'
    this.bump();
    const second = this.take();
    if (second.type !== 'num') this.fail('expected a row number after :', second.pos);
    const a = parseRowOnly(first.text);
    const b = parseRowOnly(second.text);
    if (!a || !b) this.fail(`not a valid row range: ${first.text}:${second.text}`, first.pos);
    return this.makeRange(a, b, sheet, first);
  }

  private makeRange(a: CellRef, b: CellRef, sheet: string | undefined, at: Token): Node {
    const aWhole = a.colOnly || a.rowOnly;
    const bWhole = b.colOnly || b.rowOnly;
    if (Boolean(aWhole) !== Boolean(bWhole) || a.colOnly !== b.colOnly || a.rowOnly !== b.rowOnly) {
      this.fail('mismatched range endpoints', at.pos);
    }
    // Normalise so top-left comes first; Excel accepts B2:A1.
    const lo: CellRef = { ...a };
    const hi: CellRef = { ...b };
    if (!a.rowOnly && lo.col > hi.col) {
      const t = lo.col;
      lo.col = hi.col;
      hi.col = t;
    }
    if (!a.colOnly && lo.row > hi.row) {
      const t = lo.row;
      lo.row = hi.row;
      hi.row = t;
    }
    if (a.colOnly) {
      lo.row = 1;
      hi.row = MAX_ROWS;
    }
    if (a.rowOnly) {
      lo.col = 1;
      hi.col = MAX_COLS;
    }
    return { k: 'ref', sheet, a: lo, b: hi };
  }

  private parseArgs(): Node[] {
    const args: Node[] = [];
    if (this.peek().type === 'rparen') {
      this.bump();
      return args;
    }
    for (;;) {
      // Omitted arguments are legal and meaningful: IF(A1,,2), OFFSET(A1,1,).
      const t = this.peek();
      if (t.type === 'comma' || t.type === 'rparen') args.push({ k: 'name', name: ' missing' });
      else args.push(this.parseExpr(0));

      const sep = this.take();
      if (sep.type === 'rparen') return args;
      if (sep.type !== 'comma') this.fail('expected , or ) in argument list', sep.pos);
    }
  }

  private parseArrayConstant(open: Token): Node {
    const rows: Node[][] = [];
    let row: Node[] = [];
    for (;;) {
      row.push(this.parseArrayElement());
      const sep = this.take();
      if (sep.type === 'comma') continue;
      if (sep.type === 'semi') {
        rows.push(row);
        row = [];
        continue;
      }
      if (sep.type === 'rbrace') {
        rows.push(row);
        break;
      }
      this.fail('expected , ; or } in array constant', sep.pos);
    }
    const width = rows[0]!.length;
    if (rows.some((r) => r.length !== width)) this.fail('ragged array constant', open.pos);
    return { k: 'array', rows };
  }

  private parseArrayElement(): Node {
    let sign = 1;
    let t = this.take();
    while (t.type === 'op' && (t.text === '-' || t.text === '+')) {
      if (t.text === '-') sign = -sign;
      t = this.take();
    }
    switch (t.type) {
      case 'num':
        return { k: 'num', v: sign * Number(t.text) };
      case 'str':
        return { k: 'str', v: t.text };
      case 'err':
        return { k: 'err', v: errorOf(t.text) ?? ERR.value };
      case 'ident':
        if (/^(TRUE|FALSE)$/i.test(t.text)) return { k: 'bool', v: t.text.toUpperCase() === 'TRUE' };
        this.fail('array constants may only contain literals', t.pos);
        break;
      default:
        this.fail('array constants may only contain literals', t.pos);
    }
  }
}

function cornerOf(text: string): CellRef | undefined {
  return parseCellRef(text) ?? parseColOnly(text) ?? parseRowOnly(text);
}

export function normalizeFnName(raw: string): string {
  let name = raw.toUpperCase();
  for (const p of FN_PREFIXES) if (name.startsWith(p)) name = name.slice(p.length);
  return name;
}

/** The sentinel an omitted argument parses to — `IF(A1,,2)`. */
export const MISSING_ARG = ' missing';
export const isMissing = (n: Node): boolean => n.k === 'name' && n.name === MISSING_ARG;

/**
 * Parse formula text (with or without a leading `=`) into an AST.
 * Throws `ParseError` for malformed text and `Unsupported` for shapes we
 * deliberately refuse — the caller turns either into a ⚠ cell, never a value.
 */
export function parseFormula(src: string): Node {
  let text = src.trim();
  if (text.startsWith('=')) text = text.slice(1);
  // Legacy array-formula braces: {=SUM(A1:A3*B1:B3)}
  if (text.startsWith('{') && text.endsWith('}') && text.includes('=')) {
    throw new Unsupported('FEATURE', 'legacy array formulas ({=…}) are not supported');
  }
  if (text.trim() === '') throw new ParseError('empty formula', 0, src);
  return new Parser(tokenize(text), text).parse();
}
