import { Fragment, type CSSProperties, type ReactNode } from 'react';
import type ExcelJS from 'exceljs';
import type { PreviewModel } from '../bind.js';
import { plainText, resolveContent, type CellContent } from './format.js';
import { resolveColor } from './theme.js';

/**
 * The renderer: one ExcelJS worksheet plus the engine's value overlay, painted
 * as a styled HTML table. Number formats, theme and indexed colours, fills,
 * fonts, borders, alignment, merged cells, frozen-pane sticky headers, rich text
 * and hyperlinks all come from the styled parse; every *value* comes from the
 * overlay.
 *
 * The provenance of each value is visible rather than implied:
 *   computed    a thin left edge tint, off by default
 *   cached      rendered plainly — it is the file's own number
 *   unsupported ⚠ with the reason on hover. Never a number.
 *   mismatch    outlined when diff mode is on: the file's cached value and our
 *               computed one disagree, which in a generated model usually means
 *               a total was hardcoded rather than summed.
 */

const BORDER_W: Record<string, string> = {
  hair: '1px solid',
  thin: '1px solid',
  dotted: '1px dotted',
  dashed: '1px dashed',
  medium: '2px solid',
  thick: '3px solid',
  double: '3px double',
};

function borderCss(side: unknown): string | undefined {
  if (!side || typeof side !== 'object') return undefined;
  const s = side as { style?: string; color?: unknown };
  if (!s.style) return undefined;
  return `${BORDER_W[s.style] ?? '1px solid'} ${resolveColor(s.color) ?? '#c9d2e0'}`;
}

const colPx = (w?: number, hidden?: boolean): number => (hidden ? 0 : Math.round((w ?? 8.43) * 7 + 5));
const rowPx = (h?: number, hidden?: boolean): number => (hidden ? 0 : Math.round((h ?? 15) * (96 / 72)));

function decodeRange(range: string): { l: number; t: number; r: number; b: number } {
  const [a, b] = range.split(':');
  const m1 = /([A-Z]+)(\d+)/.exec(a!)!;
  const m2 = /([A-Z]+)(\d+)/.exec(b ?? a!)!;
  const col = (s: string): number => s.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
  return { l: col(m1[1]!), t: Number(m1[2]), r: col(m2[1]!), b: Number(m2[2]) };
}

export interface ExcelViewProps {
  worksheet: ExcelJS.Worksheet;
  model: PreviewModel;
  /** Index of this sheet in the engine — matched by name at load time. */
  sheetIndex: number;
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
}

export default function ExcelView({
  worksheet: ws,
  model,
  sheetIndex,
  gridlines = false,
  zoom = 1,
  showProvenance = false,
  diffMode = false,
  flagged,
  highlight,
  onSelect,
}: ExcelViewProps): ReactNode {
  const nCols = Math.max(1, ws.columnCount);
  const nRows = Math.max(1, ws.rowCount);

  const widths: number[] = [0];
  for (let c = 1; c <= nCols; c++) {
    const col = ws.getColumn(c);
    widths.push(colPx(col.width, col.hidden));
  }
  const heights: number[] = [0];
  for (let r = 1; r <= nRows; r++) {
    const row = ws.getRow(r);
    heights.push(rowPx(row.height, row.hidden));
  }

  // Frozen panes become sticky rows/columns, with cumulative offsets.
  const view = (ws.views ?? []).find((v) => v.state === 'frozen') as
    | { xSplit?: number; ySplit?: number }
    | undefined;
  const xSplit = view?.xSplit ?? 0;
  const ySplit = view?.ySplit ?? 0;
  const leftOff: number[] = [0, 0];
  for (let c = 1; c <= nCols; c++) leftOff[c + 1] = leftOff[c]! + widths[c]!;
  const topOff: number[] = [0, 0];
  for (let r = 1; r <= nRows; r++) topOff[r + 1] = topOff[r]! + heights[r]!;

  const spans = new Map<string, { rs: number; cs: number }>();
  const covered = new Set<string>();
  const merges = (ws.model as unknown as { merges?: string[] }).merges ?? [];
  for (const range of merges) {
    const { t, l, b, r } = decodeRange(range);
    spans.set(`${t}:${l}`, { rs: b - t + 1, cs: r - l + 1 });
    for (let rr = t; rr <= b; rr++) {
      for (let cc = l; cc <= r; cc++) if (!(rr === t && cc === l)) covered.add(`${rr}:${cc}`);
    }
  }

  const rows: ReactNode[] = [];
  for (let r = 1; r <= nRows; r++) {
    const row = ws.getRow(r);
    const cells: ReactNode[] = [];
    for (let c = 1; c <= nCols; c++) {
      if (covered.has(`${r}:${c}`)) continue;
      const styledCell = row.getCell(c);
      const overlay = model.cell(sheetIndex, r, c);
      const content = resolveContent(overlay, {
        numFmt: styledCell.numFmt,
        date1904: model.facts.date1904,
        styled: styledCell.value,
      });

      const font = (styledCell.font ?? {}) as Record<string, unknown>;
      const fill = styledCell.fill as { type?: string; pattern?: string; fgColor?: unknown } | undefined;
      const align = styledCell.alignment ?? {};
      const bd = styledCell.border ?? {};
      const span = spans.get(`${r}:${c}`);

      const bg = fill?.type === 'pattern' && fill.pattern === 'solid' ? resolveColor(fill.fgColor) : undefined;
      const stickyL = c <= xSplit;
      const stickyT = r <= ySplit;
      const g = gridlines ? '1px solid #e8ecf4' : undefined;

      const mismatched =
        diffMode &&
        ((overlay?.cached !== undefined &&
          overlay.provenance === 'computed' &&
          plainText(overlay.cached) !== plainText(overlay.value)) ||
          (flagged?.has(`${r}:${c}`) ?? false));
      const computed = showProvenance && overlay?.provenance === 'computed';
      const isHighlighted = highlight?.has(`${r}:${c}`) ?? false;

      const st: CSSProperties = {
        fontWeight: font.bold ? 700 : 400,
        fontStyle: font.italic ? 'italic' : undefined,
        textDecoration: font.underline ? 'underline' : undefined,
        fontSize: font.size ? `${font.size as number}px` : undefined,
        fontFamily: font.name ? `${font.name as string}, Calibri, sans-serif` : undefined,
        color: contentColor(content) ?? resolveColor(font.color),
        background: isHighlighted
          ? 'rgba(84,140,255,0.18)'
          : (bg ?? (stickyL || stickyT ? '#fff' : undefined)),
        textAlign:
          (align.horizontal as CSSProperties['textAlign']) ??
          (content.kind === 'text' && content.numeric ? 'right' : undefined),
        verticalAlign: align.vertical === 'top' ? 'top' : align.vertical === 'bottom' ? 'bottom' : 'middle',
        whiteSpace: align.wrapText ? 'normal' : 'nowrap',
        wordBreak: align.wrapText ? 'break-word' : undefined,
        overflow: 'hidden',
        borderTop: borderCss(bd.top) ?? g,
        borderBottom: borderCss(bd.bottom) ?? g,
        borderLeft: computed ? '2px solid rgba(70,140,90,0.55)' : (borderCss(bd.left) ?? g),
        borderRight: borderCss(bd.right) ?? g,
        outline: mismatched ? '2px solid #d94a4a' : undefined,
        outlineOffset: mismatched ? '-2px' : undefined,
        position: stickyL || stickyT ? 'sticky' : undefined,
        left: stickyL ? leftOff[c] : undefined,
        top: stickyT ? topOff[r] : undefined,
        zIndex: stickyL && stickyT ? 6 : stickyT ? 5 : stickyL ? 4 : undefined,
        cursor: onSelect ? 'cell' : undefined,
      };

      cells.push(
        <td
          key={c}
          style={st}
          rowSpan={span?.rs}
          colSpan={span?.cs}
          title={tooltip(overlay, content)}
          onClick={onSelect ? () => onSelect({ row: r, col: c }) : undefined}
        >
          {renderContent(content)}
        </td>,
      );
    }
    rows.push(
      <tr key={r} style={{ height: heights[r] }}>
        {cells}
      </tr>,
    );
  }

  const cols: ReactNode[] = [];
  for (let c = 1; c <= nCols; c++) cols.push(<col key={c} style={{ width: widths[c] }} />);

  return (
    <div style={{ transform: `scale(${zoom})`, transformOrigin: 'top left', width: 'fit-content' }}>
      <table className="xl-table">
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

function tooltip(overlay: ReturnType<PreviewModel['cell']>, content: CellContent): string | undefined {
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
      return (
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
