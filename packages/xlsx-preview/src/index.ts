/**
 * @xlscalc/xlsx-preview — a client-side, view-only `.xlsx` preview that computes
 * formula values in the browser.
 *
 * The problem it exists for: an agent writing a workbook emits formulas with an
 * empty value cache (`<f>SUM(B4:E4)</f><v></v>`). Excel computes on open; every
 * other consumer reads the cache and shows nothing, so a generated model renders
 * as a labelled skeleton. This package computes those values locally — no
 * server, nothing uploaded — and marks anything it could not compute rather than
 * guessing.
 *
 * Nothing exported here needs React. The React renderer lives at
 * `@xlscalc/xlsx-preview/view`, and `renderToHtml` below serves everyone else.
 */

export { loadXlsx, type LoadOptions, type PreviewDocument, type SheetInfo } from './parse.js';
export { inspectXlsx, type Inspection, type InspectOptions } from './inspect.js';
export {
  bind,
  type BindOptions,
  type FileFacts,
  type OverlayCell,
  type PreviewModel,
  type RenderModel,
  type RenderSource,
} from './bind.js';
export { findHardcoded, describeFinding, type HardcodedFinding } from './audit.js';
export { readXlsx, type RawCell, type RawSheet, type RawWorkbook } from './ooxml.js';
export { parseXml, type XmlEvent } from './xml.js';

export {
  layoutSheet,
  blankLayout,
  layoutKey,
  mergeMap,
  paneOffsets,
  isSafeHref,
  renderExtent,
  truncationNote,
  DEFAULT_RENDER_LIMITS,
  type RenderExtent,
  type RenderLimits,
  type MergeMap,
  type BorderStyle,
  type CellAlign,
  type CellBorder,
  type CellContentData,
  type CellFont,
  type CellStyle,
  type LayoutOptions,
  type MergeRange,
  type RichRun,
  type SheetLayout,
} from './layout.js';

export {
  resolveContent,
  fromScalar,
  plainText,
  type CellContent,
  type ResolveOptions,
} from './format.js';
export { cellCss, cssText, borderCss, type CellCssOptions, type CssDecls } from './css.js';
export { resolveColor } from './theme.js';
export { renderToHtml, type RenderHtmlOptions } from './html.js';

export {
  toSnapshot,
  fromSnapshot,
  encodeScalar,
  decodeScalar,
  type OverlayCellSnapshot,
  type PreviewSnapshot,
  type RestoredDocument,
  type ScalarSnapshot,
} from './snapshot.js';

export type { EvalReport, EvalStats, Gap, Provenance, Scalar } from '@xlscalc/formula-engine';
