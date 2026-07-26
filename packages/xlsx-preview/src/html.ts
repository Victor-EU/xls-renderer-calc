/**
 * The same grid, as an HTML string, for everyone who is not using React.
 *
 * A preview is wanted in places React is not: a server rendering a thumbnail, a
 * Vue or Svelte host, an email, a PDF pipeline, a test that wants to assert on
 * markup. The React component is already a plain table, so serving those costs
 * one file rather than a second renderer — and because both go through
 * `cellCss`, the two cannot drift into disagreeing about a cell.
 *
 * The stylesheet is *not* inlined by default, and that is deliberate: shipping a
 * copy of it in this module would make two sources of truth for the one file
 * that decides whether a refused cell is visible. Read the real one and pass it
 * in when you want a standalone document:
 *
 *     import { readFileSync } from 'node:fs';
 *     const css = readFileSync(
 *       new URL(import.meta.resolve('@xlscalc/xlsx-preview/view/style.css')),
 *       'utf8',
 *     );
 *     const page = renderToHtml(doc, 0, { document: true, css });
 */

import { cellCss, cssText } from './css.js';
import { plainText, resolveContent, type CellContent } from './format.js';
import { layoutKey, mergeMap, paneOffsets, type CellStyle } from './layout.js';
import type { OverlayCell, RenderSource } from './bind.js';

const EMPTY_STYLE: CellStyle = {};

export interface RenderHtmlOptions {
  /** Draw a faint rule on every edge the file did not style itself. */
  gridlines?: boolean;
  /** Mark cells the engine computed, rather than read from the file. */
  showProvenance?: boolean;
  /** Outline cells where the file's cached value disagrees with the computed one. */
  diffMode?: boolean;
  /** Extra cells to outline in diff mode, keyed `row:col` — see audit.ts. */
  flagged?: ReadonlySet<string>;
  /** Emit `title` attributes carrying the formula and provenance. On by default. */
  tooltips?: boolean;
  /** Wrap the table in a complete HTML document. */
  document?: boolean;
  /** Stylesheet text to inline, when `document` is set. */
  css?: string;
  /** Document title, when `document` is set. Defaults to the sheet name. */
  title?: string;
}

/** Render one sheet of a loaded document as HTML. */
export function renderToHtml(doc: RenderSource, sheet: number, opts: RenderHtmlOptions = {}): string {
  const layout = doc.layouts[sheet];
  if (!layout) return '';
  const model = doc.model;
  const merges = mergeMap(layout);
  const offsets = paneOffsets(layout);
  const tooltips = opts.tooltips !== false;

  const out: string[] = [];
  out.push('<table class="xl-table"><colgroup>');
  for (let c = 1; c <= layout.cols; c++) {
    out.push(`<col style="width:${layout.colWidths[c] ?? 0}px">`);
  }
  out.push('</colgroup><tbody>');

  for (let r = 1; r <= layout.rows; r++) {
    out.push(`<tr style="height:${layout.rowHeights[r] ?? 0}px">`);
    for (let c = 1; c <= layout.cols; c++) {
      if (merges.covered(r, c)) continue;
      const key = layoutKey(r, c);
      const style = layout.styles[layout.styleAt.get(key) ?? 0] ?? EMPTY_STYLE;
      const overlay = model.cell(sheet, r, c);
      const content = resolveContent(overlay, {
        numFmt: style.numFmt,
        date1904: model.facts.date1904,
        content: layout.content.get(key),
      });

      const marks: string[] = [];
      if (opts.showProvenance && overlay?.provenance === 'computed') marks.push('xl-computed');
      if (opts.diffMode && isMismatch(overlay, opts.flagged, r, c)) marks.push('xl-mismatch');

      const color = contentColor(content);
      const css = cssText(
        cellCss(style, {
          ...(opts.gridlines === undefined ? {} : { gridlines: opts.gridlines }),
          numeric: content.kind === 'text' && content.numeric === true,
          ...(color === undefined ? {} : { color }),
          ...(c <= layout.frozenCols ? { stickyLeft: offsets.left[c]! } : {}),
          ...(r <= layout.frozenRows ? { stickyTop: offsets.top[r]! } : {}),
        }),
      );

      const span = merges.span(r, c);
      const attrs = [
        marks.length > 0 ? ` class="${marks.join(' ')}"` : '',
        css ? ` style="${esc(css)}"` : '',
        span && span.rows > 1 ? ` rowspan="${span.rows}"` : '',
        span && span.cols > 1 ? ` colspan="${span.cols}"` : '',
      ];
      if (tooltips) {
        const tip = tooltip(overlay, content);
        if (tip) attrs.push(` title="${esc(tip)}"`);
      }
      out.push(`<td${attrs.join('')}>${contentHtml(content)}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</tbody></table>');

  const table = out.join('');
  if (!opts.document) return table;

  const title = esc(opts.title ?? layout.name);
  const css = opts.css ? `<style>${opts.css}</style>` : '';
  return `<!doctype html><html><head><meta charset="utf-8"><title>${title}</title>${css}</head><body>${table}</body></html>`;
}

function isMismatch(
  overlay: OverlayCell | undefined,
  flagged: ReadonlySet<string> | undefined,
  row: number,
  col: number,
): boolean {
  if (flagged?.has(`${row}:${col}`)) return true;
  if (!overlay || overlay.cached === undefined || overlay.provenance !== 'computed') return false;
  return plainText(overlay.cached) !== plainText(overlay.value);
}

function contentColor(c: CellContent): string | undefined {
  if (c.kind === 'text') return c.color;
  if (c.kind === 'error') return '#c0392b';
  return undefined;
}

function tooltip(overlay: OverlayCell | undefined, content: CellContent): string | undefined {
  if (!overlay) return undefined;
  const lines: string[] = [];
  if (overlay.formula) lines.push(`=${overlay.formula.replace(/^=/, '')}`);
  if (content.kind === 'unsupported') lines.push(`⚠ ${content.reason}`);
  else lines.push(`${overlay.provenance}: ${plainText(overlay.value)}`);
  if (overlay.cached !== undefined && plainText(overlay.cached) !== plainText(overlay.value)) {
    lines.push(`file said: ${plainText(overlay.cached)}`);
  }
  return lines.join('\n');
}

function contentHtml(c: CellContent): string {
  switch (c.kind) {
    case 'empty':
      return '';
    case 'unsupported':
      return `<span class="xl-unsupported" title="${esc(c.reason)}">⚠</span>`;
    case 'link':
      // `noreferrer` and an escaped href: the URL comes out of someone else's
      // file, so it is untrusted input on the way to an anchor.
      return `<a href="${esc(c.href)}" target="_blank" rel="noreferrer">${esc(c.text)}</a>`;
    case 'rich':
      return c.runs
        .map((run) => {
          const style = cssText({
            fontWeight: run.bold ? 700 : undefined,
            fontStyle: run.italic ? 'italic' : undefined,
            fontSize: run.size ? `${run.size}px` : undefined,
            fontFamily: run.name ? `${run.name}, Calibri, sans-serif` : undefined,
            color: run.color,
          });
          return `<span${style ? ` style="${esc(style)}"` : ''}>${esc(run.text)}</span>`;
        })
        .join('');
    default:
      return esc(c.text);
  }
}

/**
 * Escape for both text and attribute positions.
 *
 * Every string here — cell text, hyperlink targets, sheet names, the formula in
 * a tooltip — came out of a file this library did not write. A viewer that
 * interpolates that into markup unescaped is an XSS hole with a spreadsheet in
 * front of it, so this is applied at every insertion point without exception.
 */
const ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

const esc = (s: string): string => s.replace(/[&<>"']/g, (ch) => ESCAPES[ch]!);
