import { Fragment, type CSSProperties, type ReactNode } from 'react';
import { cellCss } from '../css.js';
import { plainText, resolveContent, type CellContent } from '../format.js';
import {
  layoutKey,
  mergeMap,
  paneOffsets,
  renderExtent,
  truncationNote,
  type CellStyle,
  type RenderLimits,
} from '../layout.js';
import type { OverlayCell, RenderModel, RenderSource } from '../bind.js';

/**
 * The renderer: a sheet's layout plus the engine's value overlay, painted as a
 * styled HTML table. Number formats, colours, fills, fonts, borders, alignment,
 * merged cells, frozen-pane sticky headers, rich text and hyperlinks all come
 * from the layout; every *value* comes from the overlay.
 *
 * It takes a loaded document and a sheet index, and nothing else is required.
 * It used to take an `ExcelJS.Worksheet`, which meant every caller installed
 * ExcelJS, imported its types, and had to know that the worksheet must be
 * looked up by name with an index fallback because the two orders can disagree.
 * That was our implementation detail leaking into their call site; see
 * layout.ts.
 *
 * The provenance of each value is visible rather than implied:
 *   computed    a thin left edge tint, off by default
 *   cached      rendered plainly — it is the file's own number
 *   unsupported ⚠ with the reason on hover. Never a number.
 *   mismatch    outlined when diff mode is on: the file's cached value and our
 *               computed one disagree, which in a generated model usually means
 *               a total was hardcoded rather than summed.
 *
 * The marks are class names, not inline styles, so a host can restyle them —
 * see view/style.css, which must be imported for any of this to be visible.
 */

const EMPTY_STYLE: CellStyle = {};

export interface ExcelViewProps {
  /** A loaded document: `loadXlsx(…)`, or the result of the worker path. */
  doc: RenderSource;
  /** Which sheet to render, by index into `doc.layouts`. */
  sheet: number;
  gridlines?: boolean;
  zoom?: number;
  /** Tint cells we computed, so it is obvious which numbers were not in the file. */
  showProvenance?: boolean;
  /** Outline cells where the file's cached value disagrees with ours. */
  diffMode?: boolean;
  /**
   * Extra cells to outline in diff mode, keyed `row:col` — the hardcoded ones,
   * which have no formula of their own to disagree with and so are found by the
   * shape of their neighbourhood instead (see audit.ts).
   */
  flagged?: ReadonlySet<string>;
  /** Cells to highlight as precedents of the selected cell. */
  highlight?: ReadonlySet<string>;
  onSelect?: (address: { row: number; col: number }) => void;
  /**
   * How much of the sheet to draw. Every cell of the extent becomes a `<td>`, so
   * a sheet past this is cut short and captioned rather than left to freeze the
   * tab — see `DEFAULT_RENDER_LIMITS`.
   */
  limits?: Partial<RenderLimits>;
}

export default function ExcelView({
  doc,
  sheet: sheetIndex,
  gridlines = false,
  zoom = 1,
  showProvenance = false,
  diffMode = false,
  flagged,
  highlight,
  onSelect,
  limits,
}: ExcelViewProps): ReactNode {
  const layout = doc.layouts[sheetIndex];
  if (!layout) return null;
  const model = doc.model;

  const { colWidths, rowHeights } = layout;
  const extent = renderExtent(layout, limits);
  const { rows: nRows, cols: nCols } = extent;
  const merges = mergeMap(layout);
  const offsets = paneOffsets(layout, extent);
  const cut = truncationNote(layout, extent);

  const rows: ReactNode[] = [];
  for (let r = 1; r <= nRows; r++) {
    const cells: ReactNode[] = [];
    for (let c = 1; c <= nCols; c++) {
      if (merges.covered(r, c)) continue;
      const key = layoutKey(r, c);
      const style = layout.styles[layout.styleAt.get(key) ?? 0] ?? EMPTY_STYLE;
      const overlay = model.cell(sheetIndex, r, c);
      const content = resolveContent(overlay, {
        numFmt: style.numFmt,
        date1904: model.facts.date1904,
        content: layout.content.get(key),
      });

      const mismatched =
        diffMode &&
        ((overlay?.cached !== undefined &&
          overlay.provenance === 'computed' &&
          plainText(overlay.cached) !== plainText(overlay.value)) ||
          (flagged?.has(`${r}:${c}`) ?? false));

      const marks: string[] = [];
      if (showProvenance && overlay?.provenance === 'computed') marks.push('xl-computed');
      if (mismatched) marks.push('xl-mismatch');
      if (highlight?.has(`${r}:${c}`)) marks.push('xl-highlight');

      // Clamped to the drawn extent: a span reaching past the last column makes
      // the browser invent columns, and everything below it slides sideways.
      const span = merges.span(r, c);
      const spanRows = span ? Math.min(span.rows, nRows - r + 1) : 1;
      const spanCols = span ? Math.min(span.cols, nCols - c + 1) : 1;
      const color = contentColor(content);
      const css = cellCss(style, {
        gridlines,
        numeric: content.kind === 'text' && content.numeric === true,
        ...(color === undefined ? {} : { color }),
        ...(c <= layout.frozenCols ? { stickyLeft: offsets.left[c]! } : {}),
        ...(r <= layout.frozenRows ? { stickyTop: offsets.top[r]! } : {}),
        selectable: onSelect !== undefined,
      }) as CSSProperties;

      cells.push(
        <td
          key={c}
          className={marks.length > 0 ? marks.join(' ') : undefined}
          style={css}
          rowSpan={spanRows > 1 ? spanRows : undefined}
          colSpan={spanCols > 1 ? spanCols : undefined}
          title={tooltip(overlay, content)}
          onClick={onSelect ? () => onSelect({ row: r, col: c }) : undefined}
        >
          {renderContent(content)}
        </td>,
      );
    }
    rows.push(
      <tr key={r} style={{ height: rowHeights[r] }}>
        {cells}
      </tr>,
    );
  }

  const cols: ReactNode[] = [];
  for (let c = 1; c <= nCols; c++) cols.push(<col key={c} style={{ width: colWidths[c] }} />);

  return (
    <div style={{ transform: `scale(${zoom})`, transformOrigin: 'top left', width: 'fit-content' }}>
      <table className="xl-table">
        {cut ? <caption className="xl-truncated">{cut}</caption> : null}
        <colgroup>{cols}</colgroup>
        <tbody>{rows}</tbody>
      </table>
    </div>
  );
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

function renderContent(c: CellContent): ReactNode {
  switch (c.kind) {
    case 'empty':
      return null;
    case 'unsupported':
      return (
        <span className="xl-unsupported" title={c.reason}>
          ⚠
        </span>
      );
    case 'link':
      // A `javascript:` or `data:` target out of someone else's workbook is code,
      // not a link. The text still renders, marked, with the target on hover —
      // see `isSafeHref`.
      return c.unsafe ? (
        <span className="xl-blocked-link" title={`link not followed — ${c.href}`}>
          {c.text}
        </span>
      ) : (
        <a href={c.href} target="_blank" rel="noreferrer">
          {c.text}
        </a>
      );
    case 'rich':
      return c.runs.map((run, i) => (
        <Fragment key={i}>
          <span
            style={{
              fontWeight: run.bold ? 700 : undefined,
              fontStyle: run.italic ? 'italic' : undefined,
              fontSize: run.size ? `${run.size}px` : undefined,
              fontFamily: run.name ? `${run.name}, Calibri, sans-serif` : undefined,
              color: run.color,
            }}
          >
            {run.text}
          </span>
        </Fragment>
      ));
    default:
      return c.text;
  }
}

export type { RenderModel, RenderSource };
