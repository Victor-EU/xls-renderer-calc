// Resolve any ExcelJS colour object to a CSS hex string.
//
// ExcelJS hands back one of several shapes:
//   { argb: 'FF0F1E3A' }          explicit
//   { theme: 4, tint: -0.25 }     theme colour + tint
//   { indexed: 10 }               legacy 56-colour palette
// Only the first is trivial. Real-world files lean heavily on theme colours, so
// a viewer that only understands argb renders half a spreadsheet in inherited
// black. We resolve theme + indexed against the default Office palette (custom
// themes embedded in xl/theme/theme1.xml are a documented follow-up).

// Default Office theme (2013+), in the order the cell `theme` attribute indexes.
// Note the well-known lt1/dk1 (0/1) and lt2/dk2 (2/3) swap vs the clrScheme order.
const THEME: string[] = [
  'FFFFFF', // 0 lt1  (background 1, white)
  '000000', // 1 dk1  (text 1, black)
  'E7E6E6', // 2 lt2  (background 2)
  '44546A', // 3 dk2  (text 2)
  '4472C4', // 4 accent1
  'ED7D31', // 5 accent2
  'A5A5A5', // 6 accent3
  'FFC000', // 7 accent4
  '5B9BD5', // 8 accent5
  '70AD47', // 9 accent6
  '0563C1', // 10 hyperlink
  '954F72', // 11 followed hyperlink
];

// Standard Excel indexed palette (the classic 56 + system entries).
const INDEXED: Record<number, string> = {
  0: '000000', 1: 'FFFFFF', 2: 'FF0000', 3: '00FF00', 4: '0000FF', 5: 'FFFF00',
  6: 'FF00FF', 7: '00FFFF', 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 11: '00FF00',
  12: '0000FF', 13: 'FFFF00', 14: 'FF00FF', 15: '00FFFF', 16: '800000', 17: '008000',
  18: '000080', 19: '808000', 20: '800080', 21: '008080', 22: 'C0C0C0', 23: '808080',
  24: '9999FF', 25: '993366', 26: 'FFFFCC', 27: 'CCFFFF', 28: '660066', 29: 'FF8080',
  30: '0066CC', 31: 'CCCCFF', 32: '000080', 33: 'FF00FF', 34: 'FFFF00', 35: '00FFFF',
  36: '800080', 37: '800000', 38: '008080', 39: '0000FF', 40: '00CCFF', 41: 'CCFFFF',
  42: 'CCFFCC', 43: 'FFFF99', 44: '99CCFF', 45: 'FF99CC', 46: 'CC99FF', 47: 'FFCC99',
  48: '3366FF', 49: '33CCCC', 50: '99CC00', 51: 'FFCC00', 52: 'FF9900', 53: 'FF6600',
  54: '666699', 55: '969696', 56: '003366', 57: '339966', 58: '003300', 59: '333300',
  60: '993300', 61: '993366', 62: '333399', 63: '333333',
  64: '000000', // system foreground
  65: 'FFFFFF', // system background
};

/** Apply an Excel tint (-1..1) to a 6-char hex, lighten (>0) or darken (<0). */
function applyTint(hex: string, tint?: number): string {
  if (!tint) return hex;
  const ch = (i: number) => parseInt(hex.slice(i, i + 2), 16);
  const adj = (v: number) => {
    const out = tint < 0 ? v * (1 + tint) : v * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(out)));
  };
  const to2 = (n: number) => n.toString(16).padStart(2, '0');
  return `${to2(adj(ch(0)))}${to2(adj(ch(2)))}${to2(adj(ch(4)))}`;
}

export function resolveColor(color: unknown): string | undefined {
  if (!color || typeof color !== 'object') return undefined;
  const c = color as { argb?: string; theme?: number; tint?: number; indexed?: number };
  if (typeof c.argb === 'string' && c.argb.length >= 6) {
    return `#${c.argb.length === 8 ? c.argb.slice(2) : c.argb}`;
  }
  if (typeof c.theme === 'number') {
    const base = THEME[c.theme];
    if (base) return `#${applyTint(base, c.tint)}`;
  }
  if (typeof c.indexed === 'number') {
    const base = INDEXED[c.indexed];
    if (base) return `#${applyTint(base, c.tint)}`;
  }
  return undefined; // unknown → inherit
}
