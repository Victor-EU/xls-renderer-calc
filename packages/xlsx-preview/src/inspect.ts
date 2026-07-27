/**
 * "What will you fail to compute in this file?" — asked before rendering it.
 *
 * This is the question a product has to answer at the routing decision, not
 * after. A file this engine covers should render in the browser; one that leans
 * on `OFFSET` and iterative calculation probably belongs on a server-side
 * LibreOffice path, or behind a warning, or nowhere. Previously the only way to
 * find out was to run the whole load and read `doc.model.report`, which meant
 * paying for the styling parse and full evaluation in order to learn that you
 * should not have.
 *
 * `inspectXlsx` reads the OOXML with this package's own reader, parses every
 * formula, and stops. No ExcelJS — so none of the 22 MB styling dependency is
 * touched — and no evaluation.
 *
 * On the largest workbook of the real corpus — 138,421 formulas — that is 0.8
 * seconds against 7.5 for the full load, and it is enough to learn that the file
 * uses `OFFSET` in 384 places, `CELL` in 24, and asks for iterative calculation.
 *
 * ## What it does not tell you
 *
 * It reports where the trouble *starts*: the 204 cells of that workbook whose
 * own formula names something the engine refuses. It does not report how many
 * cells those poison downstream, and the gap is the whole story — the engine
 * ends up refusing **64,809**, a little under half the workbook. Refusal
 * propagates by design, because a `SUM` that quietly dropped an input it could
 * not compute would be exactly the silent wrong number this project exists to
 * prevent.
 *
 * Establishing the true figure means resolving every reference and walking the
 * dependency graph, which is most of what evaluation does; a second
 * implementation of it here would be a second thing to keep in agreement with
 * the engine. So this stays honest about its scope: `unsupportedCells` is a
 * floor, never an estimate, and `loadXlsx(...).model.report.stats.unsupported`
 * is the real number.
 *
 * It is a *different* floor from the engine's own root count, and deliberately
 * the larger one. The engine stops at the first failure in a chain, so it
 * attributes only 33 of those 204 cells to `OFFSET` and `CELL` and files the
 * other 171 under "depends on a cell that could not be computed". For deciding
 * whether to render a file, "how much of this leans on what we do not support"
 * is the more useful question than "which cell noticed first".
 */

import {
  Unsupported,
  installStandardLibrary,
  isParseError,
  isUnsupported,
  parseFormula,
  refusedFunctions,
  registeredFunctions,
  walk,
} from '@xlscalc/formula-engine';
import { hasCachedValue, readXlsx } from './ooxml.js';

export interface FunctionUse {
  name: string;
  /** Call sites, not cells — one formula can name the same function twice. */
  count: number;
  /** False when the engine refuses it, or does not know it at all. */
  supported: boolean;
  /** Why it is refused, when the engine says. */
  reason?: string;
}

export interface Inspection {
  /** The application that last saved the file, when `docProps/app.xml` says. */
  writer?: string;
  sheets: Array<{ name: string; cells: number; formulas: number; uncached: number }>;
  formulas: number;
  /**
   * Formula cells the file carries no value for — these render blank elsewhere.
   * Agrees exactly with `loadXlsx(...).model.facts.uncached`: both ask
   * `hasCachedValue`, and a test holds them to the same answer.
   */
  uncached: number;
  arrayFormulas: number;
  sharedFormulas: number;
  definedNames: number;
  date1904: boolean;
  /** The workbook asks Excel to resolve its circular references by iterating. */
  iterative: boolean;
  /** Every function the file names, most-used first. */
  functions: FunctionUse[];
  /** Just the ones that will refuse. A subset of `functions`. */
  unsupported: FunctionUse[];
  /**
   * Formula cells naming at least one unsupported function. A **floor** on how
   * many cells will render ⚠, not a prediction — see the note at the top of
   * this file.
   */
  unsupportedCells: number;
  /** Formulas this engine could not parse at all. These also render ⚠. */
  parseErrors: number;
  /** True when nothing in the file is outside the engine's vocabulary. */
  fullyCovered: boolean;
  readMs: number;
}

export interface InspectOptions {
  /** Cap the number of formulas parsed. Omit to parse them all. */
  limit?: number;
}

export function inspectXlsx(buf: ArrayBuffer, opts: InspectOptions = {}): Inspection {
  installStandardLibrary();
  const t0 = now();
  const raw = readXlsx(buf);

  const known = new Set(registeredFunctions());
  const refused = new Map<string, string>();
  for (const name of refusedFunctions()) refused.set(name, '');

  const counts = new Map<string, number>();
  const reasons = new Map<string, string>();
  let formulas = 0;
  let uncached = 0;
  let arrayFormulas = 0;
  let sharedFormulas = 0;
  let unsupportedCells = 0;
  let parseErrors = 0;
  let parsed = 0;

  const sheets = raw.sheets.map((sheet) => {
    let sheetFormulas = 0;
    let sheetUncached = 0;
    for (const cell of sheet.cells) {
      if (!cell.f) continue;
      sheetFormulas++;
      formulas++;
      // The same judgement `bind` makes, from the same function. Testing `<v>`
      // alone here counted every `IF(...,"",...)` that had legitimately computed
      // to the empty string as never computed, so this reported more
      // never-computed cells than the load it is meant to predict.
      if (!hasCachedValue(cell)) {
        sheetUncached++;
        uncached++;
      }
      if (cell.arrayFormula) arrayFormulas++;
      if (cell.fromShared) sharedFormulas++;

      if (opts.limit !== undefined && parsed >= opts.limit) continue;
      parsed++;

      let bad = false;
      try {
        const ast = parseFormula(cell.f);
        walk(ast, (n) => {
          if (n.k !== 'fn') return;
          counts.set(n.name, (counts.get(n.name) ?? 0) + 1);
          if (!known.has(n.name) || refused.has(n.name)) bad = true;
        });
      } catch (e) {
        // A formula that refuses *at parse time* names its subject, which is
        // more precise than the walk above can be — record it and count the
        // cell, because it will render ⚠ either way.
        if (isUnsupported(e)) {
          bad = true;
          const subject = (e as Unsupported).subject;
          if (subject) {
            counts.set(subject, (counts.get(subject) ?? 0) + 1);
            reasons.set(subject, (e as Unsupported).message);
          }
        } else if (isParseError(e)) {
          parseErrors++;
        } else {
          throw e;
        }
      }
      if (bad) unsupportedCells++;
    }
    return {
      name: sheet.name,
      cells: sheet.cells.length,
      formulas: sheetFormulas,
      uncached: sheetUncached,
    };
  });

  const functions: FunctionUse[] = [...counts.entries()]
    .map(([name, count]) => {
      const supported = known.has(name) && !refused.has(name);
      const reason = reasons.get(name);
      return reason === undefined ? { name, count, supported } : { name, count, supported, reason };
    })
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));

  const unsupported = functions.filter((f) => !f.supported);

  return {
    ...(raw.writer === undefined ? {} : { writer: raw.writer }),
    sheets,
    formulas,
    uncached,
    arrayFormulas,
    sharedFormulas,
    definedNames: raw.definedNames.length,
    date1904: raw.date1904,
    iterative: raw.iterative,
    functions,
    unsupported,
    unsupportedCells,
    parseErrors,
    fullyCovered: unsupported.length === 0 && parseErrors === 0,
    readMs: Math.round(now() - t0),
  };
}

const now = (): number => (typeof performance !== 'undefined' ? performance.now() : Date.now());
