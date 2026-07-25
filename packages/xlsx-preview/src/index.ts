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
 */

export { loadXlsx, type LoadOptions, type PreviewDocument, type SheetInfo } from './parse.js';
export { bind, type BindOptions, type FileFacts, type OverlayCell, type PreviewModel } from './bind.js';
export { findHardcoded, describeFinding, type HardcodedFinding } from './audit.js';
export { readXlsx, type RawCell, type RawSheet, type RawWorkbook } from './ooxml.js';
export { parseXml, type XmlEvent } from './xml.js';

export type { EvalReport, EvalStats, Gap, Provenance, Scalar } from '@xlscalc/formula-engine';
