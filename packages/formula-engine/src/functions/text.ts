import { numberToText, textToNumber, toText } from '../coerce.js';
import { Unsupported } from '../errors.js';
import { scalarOf, type FnContext } from '../interpreter.js';
import { ERR, ExcelError, isErr, type EvalValue, type Scalar } from '../values.js';
import { getNumberFormatter } from './format-hook.js';
import { fn, lazyFn } from './registry.js';
import { flatten, num, optNum, trunc } from './util.js';
import { excelRound } from './math.js';

/**
 * Excel's wildcard syntax, shared by SEARCH and every criteria-taking function:
 * `*` any run, `?` one character, `~` escapes the next wildcard.
 */
export function wildcardToRegExp(pattern: string, flags = 'i'): RegExp {
  let out = '';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i]!;
    if (c === '~') {
      const nxt = pattern[i + 1];
      if (nxt === '*' || nxt === '?' || nxt === '~') {
        out += escapeRe(nxt);
        i++;
        continue;
      }
      out += escapeRe(c);
      continue;
    }
    if (c === '*') out += '[\\s\\S]*';
    else if (c === '?') out += '[\\s\\S]';
    else out += escapeRe(c);
  }
  return new RegExp(`^${out}$`, flags);
}

const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

function str(args: EvalValue[], i: number, ctx: FnContext): string | ExcelError {
  return toText(scalarOf(args[i]!, ctx));
}

export function registerText(): void {
  fn('LEN', 1, 1, (args, ctx) => {
    const s = str(args, 0, ctx);
    return isErr(s) ? s : s.length;
  });

  fn('LEFT', 1, 2, (args, ctx) => sliceText(args, ctx, 'left'));
  fn('RIGHT', 1, 2, (args, ctx) => sliceText(args, ctx, 'right'));

  fn('MID', 3, 3, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    const start = num(args, 1, ctx);
    if (isErr(start)) return start;
    const len = num(args, 2, ctx);
    if (isErr(len)) return len;
    if (start < 1 || len < 0) return ERR.value;
    return s.slice(trunc(start) - 1, trunc(start) - 1 + trunc(len));
  });

  fn('UPPER', 1, 1, (args, ctx) => mapText(args, ctx, (s) => s.toUpperCase()));
  fn('LOWER', 1, 1, (args, ctx) => mapText(args, ctx, (s) => s.toLowerCase()));
  fn('PROPER', 1, 1, (args, ctx) =>
    mapText(args, ctx, (s) => s.replace(/[A-Za-z]+/g, (w) => w[0]!.toUpperCase() + w.slice(1).toLowerCase())),
  );
  // TRIM collapses internal runs of spaces to one — it is not String.trim().
  fn('TRIM', 1, 1, (args, ctx) => mapText(args, ctx, (s) => s.replace(/ +/g, ' ').trim()));
  // CLEAN strips the ASCII control characters, U+0000 to U+001F.
  fn('CLEAN', 1, 1, (args, ctx) => mapText(args, ctx, (s) => s.replace(/[\u0000-\u001f]/g, '')));

  fn('EXACT', 2, 2, (args, ctx) => {
    const a = str(args, 0, ctx);
    if (isErr(a)) return a;
    const b = str(args, 1, ctx);
    return isErr(b) ? b : a === b; // the one text comparison that IS case-sensitive
  });

  fn('FIND', 2, 3, (args, ctx) => findImpl(args, ctx, false));
  fn('SEARCH', 2, 3, (args, ctx) => findImpl(args, ctx, true));

  fn('SUBSTITUTE', 3, 4, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    const from = str(args, 1, ctx);
    if (isErr(from)) return from;
    const to = str(args, 2, ctx);
    if (isErr(to)) return to;
    if (from === '') return s;

    const which = optNum(args, 3, ctx, 0);
    if (isErr(which)) return which;
    if (which === 0) return s.split(from).join(to);
    const n = trunc(which);
    if (n < 1) return ERR.value;

    let idx = -1;
    for (let i = 0; i < n; i++) {
      idx = s.indexOf(from, idx + (i === 0 ? 0 : from.length));
      if (idx < 0) return s;
    }
    return s.slice(0, idx) + to + s.slice(idx + from.length);
  });

  fn('REPLACE', 4, 4, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    const start = num(args, 1, ctx);
    if (isErr(start)) return start;
    const len = num(args, 2, ctx);
    if (isErr(len)) return len;
    const rep = str(args, 3, ctx);
    if (isErr(rep)) return rep;
    if (start < 1 || len < 0) return ERR.value;
    const i = trunc(start) - 1;
    return s.slice(0, i) + rep + s.slice(i + trunc(len));
  });

  fn('REPT', 2, 2, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    const n = num(args, 1, ctx);
    if (isErr(n)) return n;
    const k = trunc(n);
    if (k < 0) return ERR.value;
    if (s.length * k > 32767) return ERR.value; // Excel's cell text limit
    return s.repeat(k);
  });

  fn('CONCATENATE', 1, -1, (args, ctx) => concatScalars(args, ctx));
  fn('CONCAT', 1, -1, (args, ctx) => {
    let out = '';
    for (const { v } of flatten(args, ctx)) {
      const t = toText(v);
      if (isErr(t)) return t;
      out += t;
    }
    return out;
  });

  fn('TEXTJOIN', 3, -1, (args, ctx) => {
    const delim = str(args, 0, ctx);
    if (isErr(delim)) return delim;
    const skipEmpty = scalarOf(args[1]!, ctx);
    if (isErr(skipEmpty)) return skipEmpty;
    const skip = skipEmpty === null ? true : skipEmpty !== false && skipEmpty !== 0;
    const parts: string[] = [];
    for (const { v } of flatten(args.slice(2), ctx)) {
      if (isErr(v)) return v;
      if (skip && (v === null || v === '')) continue;
      const t = toText(v);
      if (isErr(t)) return t;
      parts.push(t);
    }
    return parts.join(delim);
  });

  fn('CHAR', 1, 1, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const k = trunc(n);
    return k < 1 || k > 255 ? ERR.value : String.fromCharCode(k);
  });
  fn('UNICHAR', 1, 1, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const k = trunc(n);
    return k < 1 || k > 0x10ffff ? ERR.value : String.fromCodePoint(k);
  });
  fn('CODE', 1, 1, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    return s.length === 0 ? ERR.value : s.charCodeAt(0);
  });
  fn('UNICODE', 1, 1, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    return s.length === 0 ? ERR.value : s.codePointAt(0)!;
  });

  fn('VALUE', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (isErr(v)) return v;
    if (typeof v === 'number') return v;
    if (v === null) return 0;
    if (typeof v === 'boolean') return ERR.value;
    const n = textToNumber(v);
    return n === undefined ? ERR.value : n;
  });

  fn('NUMBERVALUE', 1, 3, (args, ctx) => {
    const s = str(args, 0, ctx);
    if (isErr(s)) return s;
    const dec = args[1] === undefined ? '.' : str(args, 1, ctx);
    if (isErr(dec)) return dec;
    const grp = args[2] === undefined ? ',' : str(args, 2, ctx);
    if (isErr(grp)) return grp;
    const normalized = s.split(grp || ',').join('').split(dec || '.').join('.');
    const n = textToNumber(normalized);
    return n === undefined ? ERR.value : n;
  });

  fn('FIXED', 1, 3, (args, ctx) => {
    const n = num(args, 0, ctx);
    if (isErr(n)) return n;
    const d = optNum(args, 1, ctx, 2);
    if (isErr(d)) return d;
    const noCommas = args[2] === undefined ? false : scalarOf(args[2]!, ctx) === true;
    const digits = Math.max(0, trunc(d));
    const rounded = excelRound(n, trunc(d), 'half');
    const body = Math.abs(rounded).toFixed(digits);
    const [int, frac] = body.split('.');
    const grouped = noCommas ? int! : int!.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return `${rounded < 0 ? '-' : ''}${grouped}${frac ? `.${frac}` : ''}`;
  });

  /**
   * TEXT needs a real number-format implementation. Rather than ship a partial
   * one that silently mis-renders `[Red](#,##0);` sections, the engine asks the
   * host for a formatter and refuses when there isn't one (see format-hook.ts).
   */
  fn('TEXT', 2, 2, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (isErr(v)) return v;
    const code = str(args, 1, ctx);
    if (isErr(code)) return code;

    const formatter = getNumberFormatter();
    if (!formatter) {
      throw new Unsupported(
        'FEATURE',
        'TEXT() requires a number-format engine; none is installed in this host',
        'TEXT',
      );
    }
    const n = typeof v === 'number' ? v : v === null ? 0 : typeof v === 'boolean' ? (v ? 1 : 0) : textToNumber(v);
    if (n === undefined) return typeof v === 'string' ? v : ERR.value;
    try {
      return formatter(code, n);
    } catch {
      return ERR.value;
    }
  });

  // Kept lazy so an unsupported alternative branch cannot poison a good value.
  lazyFn('DOLLAR', 1, 2, (nodes, ctx, ev) => {
    const v = scalarOf(ev(nodes[0]!, ctx), ctx);
    const n = typeof v === 'number' ? v : textToNumber(String(v ?? ''));
    if (n === undefined || isErr(v)) return isErr(v) ? v : ERR.value;
    const dNode = nodes[1];
    let digits = 2;
    if (dNode) {
      const d = scalarOf(ev(dNode, ctx), ctx);
      const dn = typeof d === 'number' ? d : textToNumber(String(d ?? ''));
      if (dn === undefined) return ERR.value;
      digits = trunc(dn);
    }
    const rounded = excelRound(n, digits, 'half');
    const body = groupThousands(Math.abs(rounded).toFixed(Math.max(0, digits)));
    return rounded < 0 ? `($${body})` : `$${body}`;
  });
}

function sliceText(args: EvalValue[], ctx: FnContext, side: 'left' | 'right'): Scalar {
  const s = str(args, 0, ctx);
  if (isErr(s)) return s;
  const n = optNum(args, 1, ctx, 1);
  if (isErr(n)) return n;
  const k = trunc(n);
  if (k < 0) return ERR.value;
  return side === 'left' ? s.slice(0, k) : k === 0 ? '' : s.slice(-k);
}

function mapText(args: EvalValue[], ctx: FnContext, f: (s: string) => string): Scalar {
  const s = str(args, 0, ctx);
  return isErr(s) ? s : f(s);
}

function concatScalars(args: EvalValue[], ctx: FnContext): Scalar {
  let out = '';
  for (let i = 0; i < args.length; i++) {
    const t = toText(scalarOf(args[i]!, ctx));
    if (isErr(t)) return t;
    out += t;
  }
  return out;
}

const groupThousands = (s: string): string => {
  const [int, frac] = s.split('.');
  return int!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (frac ? '.' + frac : '');
};

function findImpl(args: EvalValue[], ctx: FnContext, wildcard: boolean): Scalar {
  const needle = str(args, 0, ctx);
  if (isErr(needle)) return needle;
  const hay = str(args, 1, ctx);
  if (isErr(hay)) return hay;
  const start = optNum(args, 2, ctx, 1);
  if (isErr(start)) return start;
  const from = trunc(start) - 1;
  if (from < 0 || from > hay.length) return ERR.value;

  if (!wildcard) {
    // FIND is case-sensitive and takes no wildcards.
    const i = hay.indexOf(needle, from);
    return i < 0 ? ERR.value : i + 1;
  }
  // SEARCH is case-insensitive and honours wildcards. The anchors come off the
  // shared pattern compiler so the match can start anywhere in the haystack.
  const body = wildcardToRegExp(needle, 'i').source.replace(/^\^/, '').replace(/\$$/, '');
  const m = new RegExp(body, 'i').exec(hay.slice(from));
  return m ? from + m.index + 1 : ERR.value;
}
