/**
 * The React renderer.
 *
 * This subpath is the only place React is required, which is why it is a subpath
 * at all: `@xlscalc/xlsx-preview` on its own loads and computes a workbook with
 * no UI framework involved.
 *
 * Importing the stylesheet is not optional — `⚠`, the mark that says a value was
 * refused rather than computed, is invisible without it:
 *
 *     import '@xlscalc/xlsx-preview/view/style.css';
 */

export { default as ExcelView, type ExcelViewProps } from './ExcelView.js';
export {
  resolveContent,
  fromScalar,
  plainText,
  type CellContent,
  type ResolveOptions,
  type RichRun,
} from '../format.js';
export { cellCss, cssText, borderCss, type CssDecls, type CellCssOptions } from '../css.js';
export { resolveColor } from '../theme.js';
