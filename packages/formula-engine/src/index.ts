/**
 * @xlscalc/formula-engine — an Excel formula parser and evaluator with no
 * dependencies and one rule: **it never renders a number it is not sure of.**
 *
 * Anything outside its capability throws `Unsupported`, the cell is marked, and
 * the reason is reported. Excel's own error values (`#DIV/0!` and friends) are
 * computed results and render as themselves — the two are never conflated.
 */

export {
  Workbook,
  Sheet,
  valuesAgree,
  type CellRecord,
  type EvalReport,
  type EvalStats,
  type FormulaCell,
  type Gap,
  type Provenance,
  type WorkbookOptions,
} from './workbook.js';

export { parseFormula, normalizeFnName } from './parser.js';
export { tokenize, type Token } from './lexer.js';
export { unparse, walk, translate, type Node, type BinOp } from './ast.js';

export {
  ERR,
  ExcelError,
  Matrix,
  RefValue,
  errorOf,
  isBlank,
  isErr,
  isMatrix,
  isRef,
  type ErrorKind,
  type EvalValue,
  type Scalar,
} from './values.js';

export {
  ParseError,
  Unsupported,
  isParseError,
  isUnsupported,
  type UnsupportedCode,
} from './errors.js';

export {
  compare,
  numberToText,
  textToNumber,
  toBool,
  toNumber,
  toText,
} from './coerce.js';

export {
  MAX_COLS,
  MAX_ROWS,
  cellId,
  colName,
  colNumber,
  formatAddress,
  parseCellRef,
  quoteSheetName,
  unpackId,
  type CellRef,
} from './a1.js';

export {
  implementedFunctions,
  refusedFunctions,
  registeredFunctions,
  setNumberFormatter,
  installStandardLibrary,
  type NumberFormatter,
} from './functions/index.js';

export {
  serialToDate,
  serialToMs,
  msToSerial,
  type DateSystem,
} from './functions/dates.js';

export { excelRound } from './functions/math.js';
