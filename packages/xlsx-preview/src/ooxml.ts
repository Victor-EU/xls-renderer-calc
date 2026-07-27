/**
 * Reading formulas straight out of the `.xlsx` zip.
 *
 * The renderer uses ExcelJS for style, layout and merges — it is good at that.
 * But ExcelJS drops the raw `t` (cell type) attribute, and that single attribute
 * is what distinguishes the two cases a preview must not confuse:
 *
 *   <c r="G9" t="str"><f>IF(G9=0,"",F9/G9-1)</f><v></v></c>   computed to ""
 *   <c r="F4"><f>SUM(B4:E4)</f><v></v></c>                     never computed
 *
 * Without `t`, both look identical, which is why the previous renderer had to
 * guess at file level with a 25%-of-formulas heuristic. Reading the part
 * directly replaces that guess with the fact.
 *
 * The same pass also recovers shared formulas properly. OOXML stores
 * `<f t="shared" ref="B2:E2" si="0">` once and every sibling references it by
 * `si`; a reader that hands back the master text verbatim reports four cells as
 * computing the first cell's numbers. Here the master's AST is translated by
 * each sibling's row/column offset.
 */

import { unzipSync } from 'fflate';
import { colNumber, parseFormula, translate, unparse, MAX_COLS, MAX_ROWS } from '@xlscalc/formula-engine';
import { parseXml } from './xml.js';

export interface RawCell {
  row: number;
  col: number;
  /** The raw `t` attribute: n | s | str | b | e | inlineStr | d. Absent for plain numbers. */
  t?: string;
  /** Formula text, with shared-formula offsets already applied. */
  f?: string;
  /** Set when the formula was reconstructed from a shared master. */
  fromShared?: boolean;
  /** Set for array-formula masters, which spill and are reported as a gap. */
  arrayFormula?: boolean;
  /** Raw `<v>` text. `''` when the element was present but empty; undefined when absent. */
  v?: string;
  /** Inline string content, for `t="inlineStr"`. */
  is?: string;
}

export interface RawSheet {
  name: string;
  cells: RawCell[];
}

export interface RawWorkbook {
  sheets: RawSheet[];
  sharedStrings: string[];
  definedNames: Array<{ name: string; sheet?: number; formula: string }>;
  date1904: boolean;
  /** openpyxl sets this; LibreOffice strips it. Reported, not acted on. */
  fullCalcOnLoad: boolean;
  /** A stale calcChain makes Excel complain, so write-back must not author one. */
  hasCalcChain: boolean;
  /**
   * The application that last saved the file, from `docProps/app.xml`, when it
   * says. This is not trivia: it decides how much the file's own formula cache
   * is worth. A workbook saved by Excel carries Excel's answers, which is the
   * strongest ground truth available; one exported from Google Sheets carries
   * Google's, which differ in documented ways — `=A1` on an empty cell is blank
   * rather than 0, and `LEFT` sees the displayed text rather than the value. A
   * generator that never computed anything says nothing at all.
   */
  writer?: string;
  /**
   * Set when the workbook asks for iterative calculation, which means its
   * circular references are deliberate and Excel resolves them by iterating.
   * Reported so a refusal can say "this workbook expects iteration" rather than
   * implying the model is broken.
   */
  iterative: boolean;
}

/**
 * Did the file carry a value for this cell, or has it never been computed?
 *
 * This is the judgement the header's two-line example is about, and it exists as
 * one exported function because it was being made in two places that disagreed.
 * `bind` applied the `t="str"` refinement and `inspect` did not, so a workbook
 * full of `IF(...,"",...)` — the commonest formula in a finished model — was
 * reported as having more never-computed cells by the tool you run *before*
 * loading than by the load itself. Two honest numbers for the same file, and the
 * one that arrives first is the one a router acts on.
 *
 * The rule: an absent or empty `<v>` means never computed, *unless* `t="str"`,
 * which says the formula ran and returned the empty string. Excel writes the
 * type attribute whenever it writes a result, so the attribute's presence is the
 * evidence of computation and the emptiness of `<v>` is only the result.
 *
 * Note this asks what the *file* holds, not what the cell is worth. A value
 * written by a generator that never calculated is still a value here; how much
 * to trust it is `RawWorkbook.writer`'s question, answered separately.
 */
export function hasCachedValue(cell: RawCell): boolean {
  const empty = cell.v === undefined || cell.v === '';
  return !empty || cell.t === 'str';
}

const decoder = new TextDecoder('utf-8');

export function readXlsx(buf: ArrayBuffer | Uint8Array): RawWorkbook {
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const files = unzipSync(bytes);
  const text = (path: string): string | undefined => {
    const f = files[path];
    return f ? decoder.decode(f) : undefined;
  };

  const workbookXml = text('xl/workbook.xml');
  if (!workbookXml) throw new Error('not an xlsx: xl/workbook.xml is missing');

  const rels = readRels(text('xl/_rels/workbook.xml.rels') ?? '');
  const meta = readWorkbook(workbookXml);
  const sharedStrings = readSharedStrings(text('xl/sharedStrings.xml'));

  const sheets: RawSheet[] = [];
  for (const entry of meta.sheets) {
    const target = rels[entry.rid];
    const path = target ? normalizePath(target) : `xl/worksheets/sheet${sheets.length + 1}.xml`;
    const xml = text(path);
    sheets.push({ name: entry.name, cells: xml ? readSheet(xml) : [] });
  }

  const writer = /<Application>([^<]*)<\/Application>/.exec(text('docProps/app.xml') ?? '')?.[1];

  return {
    sheets,
    sharedStrings,
    definedNames: meta.definedNames,
    date1904: meta.date1904,
    fullCalcOnLoad: meta.fullCalcOnLoad,
    hasCalcChain: Boolean(files['xl/calcChain.xml']),
    ...(writer ? { writer } : {}),
    iterative: /<calcPr[^>]*\biterate="(1|true)"/.test(workbookXml),
  };
}

function normalizePath(target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  return target.startsWith('xl/') ? target : `xl/${target}`;
}

function readRels(xml: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const e of parseXml(xml)) {
    if (e.kind === 'open' && e.name === 'Relationship' && e.attrs.Id && e.attrs.Target) {
      out[e.attrs.Id] = e.attrs.Target;
    }
  }
  return out;
}

interface WorkbookMeta {
  sheets: Array<{ name: string; rid: string }>;
  definedNames: Array<{ name: string; sheet?: number; formula: string }>;
  date1904: boolean;
  fullCalcOnLoad: boolean;
}

function readWorkbook(xml: string): WorkbookMeta {
  const meta: WorkbookMeta = { sheets: [], definedNames: [], date1904: false, fullCalcOnLoad: false };
  let pendingName: { name: string; sheet?: number } | undefined;
  let buffer = '';

  for (const e of parseXml(xml)) {
    if (e.kind === 'open') {
      switch (e.name) {
        case 'sheet':
          meta.sheets.push({ name: e.attrs.name ?? `Sheet${meta.sheets.length + 1}`, rid: e.attrs.id ?? '' });
          break;
        case 'workbookPr':
          meta.date1904 = e.attrs.date1904 === '1' || e.attrs.date1904 === 'true';
          break;
        case 'calcPr':
          meta.fullCalcOnLoad = e.attrs.fullCalcOnLoad === '1' || e.attrs.fullCalcOnLoad === 'true';
          break;
        case 'definedName': {
          const scope = e.attrs.localSheetId;
          pendingName = {
            name: e.attrs.name ?? '',
            ...(scope === undefined ? {} : { sheet: Number(scope) }),
          };
          buffer = '';
          break;
        }
        default:
          break;
      }
    } else if (e.kind === 'text' && pendingName) {
      buffer += e.text;
    } else if (e.kind === 'close' && e.name === 'definedName' && pendingName) {
      // `_xlnm.Print_Area` and friends are layout metadata, not values.
      if (pendingName.name && !pendingName.name.startsWith('_xlnm.')) {
        meta.definedNames.push({ ...pendingName, formula: buffer });
      }
      pendingName = undefined;
      buffer = '';
    }
  }
  return meta;
}

function readSharedStrings(xml: string | undefined): string[] {
  if (!xml) return [];
  const out: string[] = [];
  let current: string | undefined;
  let inText = false;

  for (const e of parseXml(xml)) {
    if (e.kind === 'open') {
      if (e.name === 'si') current = '';
      else if (e.name === 't' && current !== undefined) inText = !e.selfClosing;
    } else if (e.kind === 'text') {
      if (inText && current !== undefined) current += e.text;
    } else if (e.kind === 'close') {
      if (e.name === 't') inText = false;
      else if (e.name === 'si') {
        out.push(current ?? '');
        current = undefined;
      }
    }
  }
  return out;
}

interface SharedMaster {
  formula: string;
  row: number;
  col: number;
}

function readSheet(xml: string): RawCell[] {
  const cells: RawCell[] = [];
  const shared = new Map<string, SharedMaster>();

  let cell: RawCell | undefined;
  let inV = false;
  let inF = false;
  let inIs = false;
  let inIsText = false;
  let vBuf = '';
  let fBuf = '';
  let isBuf = '';
  let fAttrs: Record<string, string> = {};

  const flush = (): void => {
    if (!cell) return;
    if (inV || vBuf !== '') cell.v = vBuf;
    if (isBuf !== '') cell.is = isBuf;
    cells.push(cell);
    cell = undefined;
  };

  for (const e of parseXml(xml)) {
    if (e.kind === 'open') {
      switch (e.name) {
        case 'c': {
          const addr = e.attrs.r;
          const pos = addr ? splitAddress(addr) : undefined;
          if (!pos) {
            cell = undefined;
            break;
          }
          cell = { row: pos.row, col: pos.col };
          if (e.attrs.t) cell.t = e.attrs.t;
          vBuf = '';
          isBuf = '';
          inV = false;
          if (e.selfClosing) flush();
          break;
        }
        case 'v':
          if (!cell) break;
          // `<v/>` and `<v></v>` both mean "present but empty", which is how a
          // formula that computed to "" is stored alongside t="str".
          inV = !e.selfClosing;
          vBuf = '';
          if (e.selfClosing) cell.v = '';
          break;
        case 'f':
          if (!cell) break;
          fAttrs = e.attrs;
          fBuf = '';
          inF = !e.selfClosing;
          if (e.selfClosing) applyFormula(cell, fAttrs, '', shared);
          break;
        case 'is':
          inIs = true;
          isBuf = '';
          break;
        case 't':
          if (inIs) inIsText = !e.selfClosing;
          break;
        default:
          break;
      }
    } else if (e.kind === 'text') {
      if (inV) vBuf += e.text;
      else if (inF) fBuf += e.text;
      else if (inIsText) isBuf += e.text;
    } else if (e.kind === 'close') {
      switch (e.name) {
        case 'v':
          if (inV && cell) cell.v = vBuf;
          inV = false;
          break;
        case 'f':
          if (inF && cell) applyFormula(cell, fAttrs, fBuf, shared);
          inF = false;
          break;
        case 't':
          inIsText = false;
          break;
        case 'is':
          inIs = false;
          break;
        case 'c':
          flush();
          break;
        default:
          break;
      }
    }
  }
  flush();
  return cells;
}

function applyFormula(
  cell: RawCell,
  attrs: Record<string, string>,
  body: string,
  shared: Map<string, SharedMaster>,
): void {
  const kind = attrs.t;
  if (kind === 'array') cell.arrayFormula = true;

  if (kind === 'shared') {
    const si = attrs.si ?? '';
    if (body.trim() !== '') {
      // The master carries the text; siblings carry only the si.
      shared.set(si, { formula: body, row: cell.row, col: cell.col });
      cell.f = body;
      return;
    }
    const master = shared.get(si);
    if (!master) return; // no master seen: the cell reports itself uncomputable
    cell.f = translateFormula(master.formula, cell.row - master.row, cell.col - master.col);
    cell.fromShared = true;
    return;
  }

  if (body.trim() !== '') cell.f = body;
}

/** Offset a master formula's relative references to a sibling's position. */
function translateFormula(master: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return master;
  try {
    return `=${unparse(translate(parseFormula(master), dRow, dCol, MAX_ROWS, MAX_COLS))}`;
  } catch {
    // If the master will not parse, hand back the original: the cell will report
    // the parse failure itself rather than silently computing the wrong offsets.
    return master;
  }
}

function splitAddress(addr: string): { row: number; col: number } | undefined {
  const m = /^([A-Za-z]{1,3})(\d{1,7})$/.exec(addr);
  if (!m) return undefined;
  const col = colNumber(m[1]!);
  if (col === 0) return undefined;
  return { row: Number(m[2]), col };
}
