/**
 * Tokenizer.
 *
 * Deliberately dumb: it emits primitives (identifier, number, string, error,
 * punctuation) and leaves reference assembly to the parser. `A1:B2` arrives as
 * three tokens, not one, because `'Income Statement'!A1:B2`, `A:A`, `1:1` and
 * `Sheet1!A1:Sheet1!B2` all differ only in how those primitives combine — and a
 * lexer that tries to recognise all five shapes at once is where mis-parses
 * that *look* like valid references come from.
 *
 * Whitespace is emitted rather than skipped: between two operands a space is
 * Excel's intersection operator (`=SUM(A1:C1 B1:B3)`), so the parser needs it.
 */

import { ParseError } from './errors.js';

export type TokenType =
  | 'num'
  | 'str'
  | 'err'
  | 'ident' // letters/digits/$/_/. — may turn out to be a ref, a name or a function
  | 'sheet' // a quoted 'sheet name' (the quotes are stripped)
  | 'bracket' // [...] — external link index or structured table reference
  | 'op'
  | 'lparen'
  | 'rparen'
  | 'lbrace'
  | 'rbrace'
  | 'comma'
  | 'semi'
  | 'bang'
  | 'at'
  | 'ws'
  | 'eof';

export interface Token {
  type: TokenType;
  /** Raw text, with string quotes and sheet quotes already unescaped. */
  text: string;
  pos: number;
}

const isDigit = (c: string) => c >= '0' && c <= '9';
// `$` is an identifier character so `$A$1` lexes as one token; `.` and `\\` so
// that `_xlfn.XLOOKUP`, `STDEV.S` and locale-legal defined names survive intact.
// Anything above U+00A0 counts as a letter (U+00A0 itself stays whitespace).
const isIdentStart = (c: string) =>
  (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_' || c === '$' || c === '\\' || c > '\u00a0';
const isIdentPart = (c: string) => isIdentStart(c) || isDigit(c) || c === '.';

const ERROR_LITERALS = [
  '#GETTING_DATA',
  '#DIV/0!',
  '#VALUE!',
  '#SPILL!',
  '#NAME?',
  '#NULL!',
  '#CALC!',
  '#REF!',
  '#NUM!',
  '#N/A',
];

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const push = (type: TokenType, text: string, pos: number) => out.push({ type, text, pos });

  while (i < src.length) {
    const c = src[i]!;
    const start = i;

    // whitespace (including the non-breaking space Excel tolerates)
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r' || c === '\u00a0') {
      while (i < src.length && ' \t\n\r\u00a0'.includes(src[i]!)) i++;
      push('ws', src.slice(start, i), start);
      continue;
    }

    // number — leading digit or `.5`
    if (isDigit(c) || (c === '.' && isDigit(src[i + 1] ?? ''))) {
      while (i < src.length && isDigit(src[i]!)) i++;
      if (src[i] === '.') {
        i++;
        while (i < src.length && isDigit(src[i]!)) i++;
      }
      // exponent, only when it is actually followed by digits ("1E" is not a number)
      if ((src[i] === 'e' || src[i] === 'E') && /^[+-]?\d/.test(src.slice(i + 1, i + 3))) {
        i++;
        if (src[i] === '+' || src[i] === '-') i++;
        while (i < src.length && isDigit(src[i]!)) i++;
      }
      push('num', src.slice(start, i), start);
      continue;
    }

    // string literal — "" is an escaped quote
    if (c === '"') {
      i++;
      let text = '';
      for (;;) {
        if (i >= src.length) throw new ParseError('unterminated string', start, src);
        if (src[i] === '"') {
          if (src[i + 1] === '"') {
            text += '"';
            i += 2;
            continue;
          }
          i++;
          break;
        }
        text += src[i];
        i++;
      }
      push('str', text, start);
      continue;
    }

    // quoted sheet name — 'Income Statement'!A1
    if (c === "'") {
      i++;
      let text = '';
      for (;;) {
        if (i >= src.length) throw new ParseError('unterminated sheet name', start, src);
        if (src[i] === "'") {
          if (src[i + 1] === "'") {
            text += "'";
            i += 2;
            continue;
          }
          i++;
          break;
        }
        text += src[i];
        i++;
      }
      push('sheet', text, start);
      continue;
    }

    // error literal
    if (c === '#') {
      const rest = src.slice(i).toUpperCase();
      const lit = ERROR_LITERALS.find((e) => rest.startsWith(e));
      if (!lit) throw new ParseError('unrecognised error literal', i, src);
      i += lit.length;
      push('err', lit, start);
      continue;
    }

    // bracketed run — [1]Sheet1!A1 (external) or Table1[Column] (structured)
    if (c === '[') {
      let depth = 0;
      for (;;) {
        if (i >= src.length) throw new ParseError('unterminated [', start, src);
        if (src[i] === '[') depth++;
        else if (src[i] === ']' && --depth === 0) {
          i++;
          break;
        }
        i++;
      }
      push('bracket', src.slice(start, i), start);
      continue;
    }

    if (isIdentStart(c)) {
      i++;
      while (i < src.length && isIdentPart(src[i]!)) i++;
      push('ident', src.slice(start, i), start);
      continue;
    }

    switch (c) {
      case '(':
        i++;
        push('lparen', '(', start);
        continue;
      case ')':
        i++;
        push('rparen', ')', start);
        continue;
      case '{':
        i++;
        push('lbrace', '{', start);
        continue;
      case '}':
        i++;
        push('rbrace', '}', start);
        continue;
      case ',':
        i++;
        push('comma', ',', start);
        continue;
      case ';':
        i++;
        push('semi', ';', start);
        continue;
      case '!':
        i++;
        push('bang', '!', start);
        continue;
      case '@':
        i++;
        push('at', '@', start);
        continue;
      case '<':
        if (src[i + 1] === '=' || src[i + 1] === '>') {
          push('op', src.slice(i, i + 2), start);
          i += 2;
        } else {
          i++;
          push('op', '<', start);
        }
        continue;
      case '>':
        if (src[i + 1] === '=') {
          push('op', '>=', start);
          i += 2;
        } else {
          i++;
          push('op', '>', start);
        }
        continue;
      case '=':
      case '+':
      case '-':
      case '*':
      case '/':
      case '^':
      case '&':
      case '%':
      case ':':
        i++;
        push('op', c, start);
        continue;
      default:
        throw new ParseError(`unexpected character ${JSON.stringify(c)}`, i, src);
    }
  }

  push('eof', '', src.length);
  return out;
}
