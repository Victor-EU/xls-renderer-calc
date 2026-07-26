/**
 * Dates are serial numbers; time is the fractional part.
 *
 * **The 1900 leap-year bug is reproduced deliberately.** Serial 60 is
 * 1900-02-29 — a day that never existed. Lotus 1-2-3 had the bug, Excel copied
 * it for file compatibility, and every serial from 61 onward is therefore offset
 * by one from a naive calendar. An engine that "fixes" it disagrees with Excel
 * on every date in the file.
 *
 * The 1904 system (legacy Mac, `<workbookPr date1904="1"/>`) is a different
 * epoch with no phantom day; it is supported rather than refused, because
 * getting it wrong is a silent four-year error.
 */

import { textToNumber } from '../coerce.js';
import { scalarOf, type FnContext } from '../interpreter.js';
import { ERR, ExcelError, isErr, type EvalValue, type Scalar } from '../values.js';
import { fn } from './registry.js';
import { asMatrix, num, optNum, trunc } from './util.js';

const MS_PER_DAY = 86_400_000;
/** Serial of 1970-01-01 in the 1900 system (phantom leap day included). */
const EPOCH_1900 = 25569;
/** Serial of 1970-01-01 in the 1904 system. */
const EPOCH_1904 = 24107;

export interface DateSystem {
  readonly date1904: boolean;
}

/** Serial → UTC milliseconds. */
export function serialToMs(serial: number, sys: DateSystem): number {
  if (sys.date1904) return Math.round((serial - EPOCH_1904) * MS_PER_DAY);
  // Serials below 60 predate the phantom 1900-02-29 and need a day added back.
  const adjust = serial < 60 ? MS_PER_DAY : 0;
  return Math.round((serial - EPOCH_1900) * MS_PER_DAY) + adjust;
}

/** UTC milliseconds → serial. */
export function msToSerial(ms: number, sys: DateSystem): number {
  if (sys.date1904) return ms / MS_PER_DAY + EPOCH_1904;
  const naive = ms / MS_PER_DAY + EPOCH_1900;
  return naive < 60 ? naive - 1 : naive;
}

/** Serial → a JS Date in UTC, for rendering. Never used inside evaluation. */
export function serialToDate(serial: number, sys: DateSystem): Date {
  return new Date(serialToMs(serial, sys));
}

interface Ymd {
  y: number;
  m: number; // 1-12
  d: number;
}

export function serialToYmd(serial: number, sys: DateSystem): Ymd {
  const whole = Math.floor(serial);
  // Serial 0 in the 1900 system is Excel's "January 0, 1900" — a placeholder
  // day that does not exist, sitting one before 1900-01-01 so that the epoch
  // has a zero. Excel reports YEAR 1900, MONTH 1, DAY 0 for it, and therefore
  // EOMONTH(0,0) is the 31st of January. Arithmetically the instant is
  // 1899-12-31, which is what the conversion below produces and what
  // LibreOffice and Google Sheets both report — so without this case a date
  // function lands in the wrong *month*, not merely the wrong day.
  //
  // This is not decoration: an empty cell coerces to 0, and a model still being
  // built points date functions at empty cells constantly. Two real budgets in
  // the corpus do it 184 times between them.
  if (whole === 0 && !sys.date1904) return { y: 1900, m: 1, d: 0 };
  const d = new Date(serialToMs(whole, sys));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}

export function ymdToSerial(y: number, m: number, d: number, sys: DateSystem): number {
  // Date.UTC rolls out-of-range months and days exactly as Excel's DATE does.
  const ms = Date.UTC(y, m - 1, d);
  const serial = msToSerial(ms, sys);
  return Math.round(serial);
}

function serialArg(args: EvalValue[], i: number, ctx: FnContext): number | ExcelError {
  const v = scalarOf(args[i]!, ctx);
  if (isErr(v)) return v;
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return ERR.value;
  const n = textToNumber(v);
  if (n !== undefined) return n;
  const parsed = parseDateText(v, ctx.host.date1904 ? { date1904: true } : { date1904: false });
  return parsed ?? ERR.value;
}

/** ISO and a few common written forms; anything else is #VALUE!, never a guess. */
export function parseDateText(text: string, sys: DateSystem): number | undefined {
  const t = text.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t);
  if (m) return ymdToSerial(Number(m[1]), Number(m[2]), Number(m[3]), sys);
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(t);
  if (m) return ymdToSerial(Number(m[3]), Number(m[1]), Number(m[2]), sys);
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m) return (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] ?? 0)) / 86400;
  return undefined;
}

export function registerDates(): void {
  const sysOf = (ctx: FnContext): DateSystem => ({ date1904: ctx.host.date1904 });

  fn('DATE', 3, 3, (args, ctx) => {
    const y = num(args, 0, ctx);
    if (isErr(y)) return y;
    const m = num(args, 1, ctx);
    if (isErr(m)) return m;
    const d = num(args, 2, ctx);
    if (isErr(d)) return d;
    let year = trunc(y);
    // Excel reads a year below 1900 as an offset from 1900.
    if (year >= 0 && year <= 1899) year += 1900;
    const serial = ymdToSerial(year, trunc(m), trunc(d), sysOf(ctx));
    return serial < 0 ? ERR.num : serial;
  });

  fn('TIME', 3, 3, (args, ctx) => {
    const h = num(args, 0, ctx);
    if (isErr(h)) return h;
    const mi = num(args, 1, ctx);
    if (isErr(mi)) return mi;
    const s = num(args, 2, ctx);
    if (isErr(s)) return s;
    const total = (trunc(h) * 3600 + trunc(mi) * 60 + trunc(s)) / 86400;
    return total < 0 ? ERR.num : total % 1;
  });

  fn('DATEVALUE', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (isErr(v)) return v;
    if (typeof v !== 'string') return ERR.value;
    const s = parseDateText(v, sysOf(ctx));
    return s === undefined ? ERR.value : Math.floor(s);
  });

  fn('TIMEVALUE', 1, 1, (args, ctx) => {
    const v = scalarOf(args[0]!, ctx);
    if (isErr(v)) return v;
    if (typeof v !== 'string') return ERR.value;
    const s = parseDateText(v, sysOf(ctx));
    return s === undefined ? ERR.value : s % 1;
  });

  fn('TODAY', 0, 0, (_a, ctx) => Math.floor(ctx.host.now()), { volatile: true });
  fn('NOW', 0, 0, (_a, ctx) => ctx.host.now(), { volatile: true });

  const part = (name: string, get: (ymd: Ymd) => number): void =>
    fn(name, 1, 1, (args, ctx) => {
      const s = serialArg(args, 0, ctx);
      if (isErr(s)) return s;
      if (s < 0) return ERR.num;
      return get(serialToYmd(s, sysOf(ctx)));
    });
  part('YEAR', (d) => d.y);
  part('MONTH', (d) => d.m);
  part('DAY', (d) => d.d);

  const timePart = (name: string, div: number, mod: number): void =>
    fn(name, 1, 1, (args, ctx) => {
      const s = serialArg(args, 0, ctx);
      if (isErr(s)) return s;
      if (s < 0) return ERR.num;
      const secs = Math.round((s % 1) * 86400);
      return Math.floor(secs / div) % mod;
    });
  timePart('HOUR', 3600, 24);
  timePart('MINUTE', 60, 60);
  timePart('SECOND', 1, 60);

  fn('WEEKDAY', 1, 2, (args, ctx) => {
    const s = serialArg(args, 0, ctx);
    if (isErr(s)) return s;
    const type = optNum(args, 1, ctx, 1);
    if (isErr(type)) return type;
    const dow = new Date(serialToMs(Math.floor(s), sysOf(ctx))).getUTCDay(); // 0 = Sunday
    switch (trunc(type)) {
      case 1:
        return dow + 1;
      case 2:
        return ((dow + 6) % 7) + 1;
      case 3:
        return (dow + 6) % 7;
      case 11:
        return ((dow + 6) % 7) + 1;
      case 12:
        return ((dow + 5) % 7) + 1;
      case 13:
        return ((dow + 4) % 7) + 1;
      case 14:
        return ((dow + 3) % 7) + 1;
      case 15:
        return ((dow + 2) % 7) + 1;
      case 16:
        return ((dow + 1) % 7) + 1;
      case 17:
        return dow + 1;
      default:
        return ERR.num;
    }
  });

  fn('EDATE', 2, 2, (args, ctx) => shiftMonths(args, ctx, false));
  fn('EOMONTH', 2, 2, (args, ctx) => shiftMonths(args, ctx, true));

  fn('DAYS', 2, 2, (args, ctx) => {
    const end = serialArg(args, 0, ctx);
    if (isErr(end)) return end;
    const start = serialArg(args, 1, ctx);
    if (isErr(start)) return start;
    return Math.floor(end) - Math.floor(start);
  });

  fn('DAYS360', 2, 3, (args, ctx) => {
    const start = serialArg(args, 0, ctx);
    if (isErr(start)) return start;
    const end = serialArg(args, 1, ctx);
    if (isErr(end)) return end;
    const european = args[2] === undefined ? false : scalarOf(args[2]!, ctx) === true;
    const sys = sysOf(ctx);
    return days360(serialToYmd(start, sys), serialToYmd(end, sys), european);
  });

  fn('DATEDIF', 3, 3, (args, ctx) => {
    const start = serialArg(args, 0, ctx);
    if (isErr(start)) return start;
    const end = serialArg(args, 1, ctx);
    if (isErr(end)) return end;
    const unit = scalarOf(args[2]!, ctx);
    if (isErr(unit)) return unit;
    if (typeof unit !== 'string') return ERR.value;
    if (end < start) return ERR.num;
    const sys = sysOf(ctx);
    const a = serialToYmd(start, sys);
    const b = serialToYmd(end, sys);
    switch (unit.toUpperCase()) {
      case 'D':
        return Math.floor(end) - Math.floor(start);
      case 'M':
        return completedMonths(a, b);
      case 'Y':
        return Math.floor(completedMonths(a, b) / 12);
      case 'MD':
        return b.d >= a.d ? b.d - a.d : b.d + daysInMonth(b.y, b.m - 1) - a.d;
      case 'YM':
        return completedMonths(a, b) % 12;
      case 'YD': {
        const anniversary = ymdToSerial(b.y - (afterAnniversary(a, b) ? 0 : 1), a.m, a.d, sys);
        return Math.floor(end) - anniversary;
      }
      default:
        return ERR.num;
    }
  });

  fn('YEARFRAC', 2, 3, (args, ctx) => {
    const start = serialArg(args, 0, ctx);
    if (isErr(start)) return start;
    const end = serialArg(args, 1, ctx);
    if (isErr(end)) return end;
    const basis = optNum(args, 2, ctx, 0);
    if (isErr(basis)) return basis;
    return yearFrac(start, end, trunc(basis), sysOf(ctx));
  });

  fn('NETWORKDAYS', 2, 3, (args, ctx) => {
    const start = serialArg(args, 0, ctx);
    if (isErr(start)) return start;
    const end = serialArg(args, 1, ctx);
    if (isErr(end)) return end;
    const holidays = holidaySet(args, 2, ctx);
    if (isErr(holidays)) return holidays;
    const sys = sysOf(ctx);
    const [lo, hi] = start <= end ? [start, end] : [end, start];
    let n = 0;
    for (let s = Math.floor(lo); s <= Math.floor(hi); s++) {
      const dow = new Date(serialToMs(s, sys)).getUTCDay();
      if (dow === 0 || dow === 6 || holidays.has(s)) continue;
      n++;
    }
    return start <= end ? n : -n;
  });

  fn('WORKDAY', 2, 3, (args, ctx) => {
    const start = serialArg(args, 0, ctx);
    if (isErr(start)) return start;
    const days = num(args, 1, ctx);
    if (isErr(days)) return days;
    const holidays = holidaySet(args, 2, ctx);
    if (isErr(holidays)) return holidays;
    const sys = sysOf(ctx);
    const step = days >= 0 ? 1 : -1;
    let remaining = Math.abs(trunc(days));
    let s = Math.floor(start);
    while (remaining > 0) {
      s += step;
      const dow = new Date(serialToMs(s, sys)).getUTCDay();
      if (dow === 0 || dow === 6 || holidays.has(s)) continue;
      remaining--;
    }
    return s;
  });
}

function holidaySet(args: EvalValue[], i: number, ctx: FnContext): Set<number> | ExcelError {
  const out = new Set<number>();
  if (args[i] === undefined) return out;
  const m = asMatrix(args[i]!, ctx);
  if (isErr(m)) return m;
  for (const v of m.values()) if (typeof v === 'number') out.add(Math.floor(v));
  return out;
}

function shiftMonths(args: EvalValue[], ctx: FnContext, endOfMonth: boolean): Scalar {
  const sys: DateSystem = { date1904: ctx.host.date1904 };
  const s = serialArg(args, 0, ctx);
  if (isErr(s)) return s;
  const months = num(args, 1, ctx);
  if (isErr(months)) return months;
  const { y, m, d } = serialToYmd(s, sys);
  const targetMonth = m + trunc(months);
  if (endOfMonth) {
    // Day 0 of the following month is the last day of the target month.
    return ymdToSerial(y, targetMonth + 1, 0, sys);
  }
  const clamped = Math.min(d, daysInMonth(y, targetMonth));
  return ymdToSerial(y, targetMonth, clamped, sys);
}

/** `month` is 1-based and may be out of range; Date.UTC normalises it. */
function daysInMonth(y: number, month: number): number {
  return new Date(Date.UTC(y, month, 0)).getUTCDate();
}

function completedMonths(a: Ymd, b: Ymd): number {
  let months = (b.y - a.y) * 12 + (b.m - a.m);
  if (b.d < a.d) months--;
  return months;
}

function afterAnniversary(a: Ymd, b: Ymd): boolean {
  return b.m > a.m || (b.m === a.m && b.d >= a.d);
}

function days360(a: Ymd, b: Ymd, european: boolean): number {
  let d1 = a.d;
  let d2 = b.d;
  if (european) {
    d1 = Math.min(d1, 30);
    d2 = Math.min(d2, 30);
  } else {
    if (d1 === 31 || (a.m === 2 && d1 === daysInMonth(a.y, 2))) d1 = 30;
    if (d2 === 31) d2 = d1 === 30 ? 30 : 31;
  }
  return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
}

export function yearFrac(start: number, end: number, basis: number, sys: DateSystem): Scalar {
  const [lo, hi] = start <= end ? [start, end] : [end, start];
  const a = serialToYmd(lo, sys);
  const b = serialToYmd(hi, sys);
  const days = Math.floor(hi) - Math.floor(lo);
  switch (basis) {
    case 0:
      return days360(a, b, false) / 360;
    case 1: {
      // Actual/actual. For a span of a year or less the denominator is 366 when
      // a 29 February falls inside it, otherwise 365; for longer spans it is the
      // mean length of the calendar years the span touches.
      const withinOneYear =
        a.y === b.y || (b.y === a.y + 1 && (a.m > b.m || (a.m === b.m && a.d >= b.d)));
      const denom = withinOneYear
        ? containsLeapDay(lo, hi, sys)
          ? 366
          : 365
        : yearDays(a.y, b.y) / (b.y - a.y + 1);
      return days / denom;
    }
    case 2:
      return days / 360;
    case 3:
      return days / 365;
    case 4:
      return days360(a, b, true) / 360;
    default:
      return ERR.num;
  }
}

const isLeap = (y: number): boolean => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;

/** Does a real 29 February fall between two serials, inclusive? */
function containsLeapDay(lo: number, hi: number, sys: DateSystem): boolean {
  const from = serialToYmd(lo, sys);
  const to = serialToYmd(hi, sys);
  for (let y = from.y; y <= to.y; y++) {
    if (!isLeap(y)) continue;
    const feb29 = ymdToSerial(y, 2, 29, sys);
    if (feb29 >= Math.floor(lo) && feb29 <= Math.floor(hi)) return true;
  }
  return false;
}
const yearDays = (from: number, to: number): number => {
  let n = 0;
  for (let y = from; y <= to; y++) n += isLeap(y) ? 366 : 365;
  return n;
};
