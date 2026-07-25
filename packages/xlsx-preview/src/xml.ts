/**
 * A minimal XML pull parser, ~150 lines, sufficient for the OOXML parts we read
 * directly (`worksheets/*.xml`, `sharedStrings.xml`, `workbook.xml`).
 *
 * Why not a regex, and why not DOMParser:
 *   - A regex over `<c …>…</c>` misreads inline strings and CDATA, and those
 *     appear in real files. Silently misreading a cell is the failure mode this
 *     project cares most about.
 *   - `DOMParser` builds a full node tree for a sheet that may hold a hundred
 *     thousand cells, and it does not exist in a Node consumer or a worker
 *     without a shim.
 *
 * Pull parsing keeps allocation proportional to what the caller keeps.
 */

export type XmlEvent =
  | { kind: 'open'; name: string; attrs: Record<string, string>; selfClosing: boolean }
  | { kind: 'close'; name: string }
  | { kind: 'text'; text: string };

const ENTITY: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, body: string) => {
    if (body.startsWith('#x') || body.startsWith('#X')) {
      return String.fromCodePoint(parseInt(body.slice(2), 16));
    }
    if (body.startsWith('#')) return String.fromCodePoint(parseInt(body.slice(1), 10));
    return ENTITY[body] ?? whole;
  });
}

/** Walk the document, yielding events in order. Comments and PIs are skipped. */
export function* parseXml(src: string): Generator<XmlEvent> {
  let i = 0;
  const n = src.length;

  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt < 0) {
      const tail = src.slice(i);
      if (tail.trim()) yield { kind: 'text', text: decodeEntities(tail) };
      return;
    }
    if (lt > i) {
      const text = src.slice(i, lt);
      // Whitespace between elements is not content in these parts, but
      // whitespace *inside* <t> is significant, so it is never trimmed away.
      yield { kind: 'text', text: decodeEntities(text) };
    }

    if (src.startsWith('<!--', lt)) {
      i = src.indexOf('-->', lt) + 3;
      if (i < 3) return;
      continue;
    }
    if (src.startsWith('<![CDATA[', lt)) {
      const end = src.indexOf(']]>', lt);
      yield { kind: 'text', text: src.slice(lt + 9, end < 0 ? n : end) };
      i = end < 0 ? n : end + 3;
      continue;
    }
    if (src.startsWith('<?', lt) || src.startsWith('<!', lt)) {
      i = src.indexOf('>', lt) + 1;
      if (i < 1) return;
      continue;
    }

    const gt = findTagEnd(src, lt + 1);
    if (gt < 0) return;
    const body = src.slice(lt + 1, gt);
    i = gt + 1;

    if (body.startsWith('/')) {
      yield { kind: 'close', name: stripNs(body.slice(1).trim()) };
      continue;
    }

    const selfClosing = body.endsWith('/');
    const inner = selfClosing ? body.slice(0, -1) : body;
    const spaceAt = firstSpace(inner);
    const name = stripNs(spaceAt < 0 ? inner : inner.slice(0, spaceAt));
    const attrs = spaceAt < 0 ? {} : parseAttrs(inner.slice(spaceAt + 1));
    yield { kind: 'open', name, attrs, selfClosing };
  }
}

/** `>` inside a quoted attribute value does not end the tag. */
function findTagEnd(src: string, from: number): number {
  let quote = '';
  for (let i = from; i < src.length; i++) {
    const c = src[i]!;
    if (quote) {
      if (c === quote) quote = '';
      continue;
    }
    if (c === '"' || c === "'") quote = c;
    else if (c === '>') return i;
  }
  return -1;
}

function firstSpace(s: string): number {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c === 32 || c === 9 || c === 10 || c === 13) return i;
  }
  return -1;
}

/** OOXML parts use namespace prefixes inconsistently; we key on local names. */
function stripNs(name: string): string {
  const colon = name.indexOf(':');
  return colon < 0 ? name : name.slice(colon + 1);
}

function parseAttrs(src: string): Record<string, string> {
  const out: Record<string, string> = {};
  let i = 0;
  while (i < src.length) {
    while (i < src.length && /\s/.test(src[i]!)) i++;
    const eq = src.indexOf('=', i);
    if (eq < 0) break;
    const name = stripNs(src.slice(i, eq).trim());
    let j = eq + 1;
    while (j < src.length && /\s/.test(src[j]!)) j++;
    const quote = src[j];
    if (quote !== '"' && quote !== "'") break;
    const end = src.indexOf(quote, j + 1);
    if (end < 0) break;
    out[name] = decodeEntities(src.slice(j + 1, end));
    i = end + 1;
  }
  return out;
}
