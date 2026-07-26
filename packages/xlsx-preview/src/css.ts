/**
 * Turning a `CellStyle` into CSS, once, for both renderers.
 *
 * `ExcelView` and `renderToHtml` must agree cell for cell — a React preview and
 * a server-rendered one that differ are two bugs waiting to be reported as one.
 * So the composition lives here and each renderer only decides what to do with
 * the result: React takes the object as-is, the HTML writer runs it through
 * `cssText`.
 *
 * What is deliberately *not* here: the computed edge, the mismatch outline and
 * the precedent highlight. Those are statements about the value rather than the
 * document, they have to beat the file's own inline styling, and they belong to
 * whoever is looking rather than to the cell — so they are class names in
 * `view/style.css` and a host can restyle them without touching this.
 */

import type { CellBorder, CellStyle } from './layout.js';

/** Camel-cased declarations, the shape React's `style` prop wants. */
export type CssDecls = Record<string, string | number | undefined>;

const BORDER_WIDTH: Record<string, string> = {
  hair: '1px solid',
  thin: '1px solid',
  dotted: '1px dotted',
  dashed: '1px dashed',
  medium: '2px solid',
  thick: '3px solid',
  double: '3px double',
};

export function borderCss(side: CellBorder | undefined): string | undefined {
  if (!side) return undefined;
  return `${BORDER_WIDTH[side.style] ?? '1px solid'} ${side.color ?? '#c9d2e0'}`;
}

export interface CellCssOptions {
  /** Draw a faint rule on every edge that the file did not style itself. */
  gridlines?: boolean;
  /** Numbers right-align unless the file says otherwise. */
  numeric?: boolean;
  /** Colour from the number-format code (`[Red]`) or an error; beats the font colour. */
  color?: string;
  /** Pixel offset for a frozen column, when this cell is in one. */
  stickyLeft?: number;
  /** Pixel offset for a frozen row, when this cell is in one. */
  stickyTop?: number;
  /** Whether a click handler is attached, which changes the cursor. */
  selectable?: boolean;
}

export function cellCss(style: CellStyle, opts: CellCssOptions = {}): CssDecls {
  const font = style.font ?? {};
  const align = style.align ?? {};
  const grid = opts.gridlines ? '1px solid #e8ecf4' : undefined;

  const stickyL = opts.stickyLeft !== undefined;
  const stickyT = opts.stickyTop !== undefined;

  return {
    fontWeight: font.bold ? 700 : 400,
    fontStyle: font.italic ? 'italic' : undefined,
    textDecoration: font.underline ? 'underline' : undefined,
    fontSize: font.size ? `${font.size}px` : undefined,
    fontFamily: font.name ? `${font.name}, Calibri, sans-serif` : undefined,
    color: opts.color ?? font.color,
    // A frozen cell scrolls over the ones behind it, so it needs an opaque
    // background even when the file gave it none.
    background: style.fill ?? (stickyL || stickyT ? 'var(--xl-paper, #fff)' : undefined),
    textAlign: align.horizontal ?? (opts.numeric ? 'right' : undefined),
    verticalAlign:
      align.vertical === 'top' ? 'top' : align.vertical === 'bottom' ? 'bottom' : 'middle',
    whiteSpace: align.wrap ? 'normal' : 'nowrap',
    wordBreak: align.wrap ? 'break-word' : undefined,
    overflow: 'hidden',
    borderTop: borderCss(style.top) ?? grid,
    borderBottom: borderCss(style.bottom) ?? grid,
    borderLeft: borderCss(style.left) ?? grid,
    borderRight: borderCss(style.right) ?? grid,
    position: stickyL || stickyT ? 'sticky' : undefined,
    left: stickyL ? opts.stickyLeft : undefined,
    top: stickyT ? opts.stickyTop : undefined,
    zIndex: stickyL && stickyT ? 6 : stickyT ? 5 : stickyL ? 4 : undefined,
    cursor: opts.selectable ? 'cell' : undefined,
  };
}

/** Numeric declarations that mean pixels when written as CSS text. */
const PX = new Set(['left', 'top', 'width', 'height']);

/** Serialise declarations for a `style="…"` attribute. */
export function cssText(decls: CssDecls): string {
  const out: string[] = [];
  for (const [prop, value] of Object.entries(decls)) {
    if (value === undefined || value === '') continue;
    const name = prop.replace(/[A-Z]/g, (ch) => `-${ch.toLowerCase()}`);
    out.push(`${name}:${typeof value === 'number' && PX.has(prop) ? `${value}px` : value}`);
  }
  return out.join(';');
}
