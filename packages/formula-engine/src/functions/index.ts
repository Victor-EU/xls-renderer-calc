import { registerConditional } from './conditional.js';
import { registerDates } from './dates.js';
import { registerFinancial } from './financial.js';
import { registerInfo } from './info.js';
import { registerLogical } from './logical.js';
import { registerLookup } from './lookup.js';
import { registerMath } from './math.js';
import { refuse } from './registry.js';
import { registerStats } from './stats.js';
import { registerText } from './text.js';

let installed = false;

/**
 * Register the whole library. Idempotent, so importing the engine twice in a
 * bundle does not throw on duplicate registration.
 */
export function installStandardLibrary(): void {
  if (installed) return;
  installed = true;

  registerMath();
  registerStats();
  registerLogical();
  registerText();
  registerInfo();
  registerLookup();
  registerConditional();
  registerDates();
  registerFinancial();

  // Dynamic arrays: these spill into neighbouring cells, which changes the grid
  // rather than one cell's value. A saved file already contains the spilled
  // results, so refusing here costs nothing and prevents a half-right render.
  for (const name of ['FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'TOCOL', 'TOROW']) {
    refuse(name, `${name}() spills into neighbouring cells, which this engine does not model`);
  }
  for (const name of ['MMULT', 'MINVERSE', 'MDETERM', 'FREQUENCY', 'TREND', 'GROWTH', 'LINEST', 'LOGEST']) {
    refuse(name, `${name}() returns an array whose spill this engine does not model`);
  }
  refuse('GETPIVOTDATA', 'GETPIVOTDATA() reads pivot-cache state this engine does not load');
  refuse('HYPERLINK', 'HYPERLINK() produces a link, not a computed value');
  refuse('WEBSERVICE', 'WEBSERVICE() makes a network request, which never happens in a preview');
  refuse('RTD', 'RTD() reads a live data feed');
}

export * from './registry.js';
export { setNumberFormatter, type NumberFormatter } from './format-hook.js';
