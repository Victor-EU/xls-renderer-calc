/**
 * The smallest `.xlsx` that says exactly what a test needs it to say.
 *
 * Hand-written rather than produced by a writer, because these suites are about
 * what the *bytes* claim and every writer normalises away the distinctions being
 * tested. ExcelJS will not emit `<c><f>SUM(A1:A2)</f><v></v></c>` — a formula
 * with an empty, untyped result — and that is the exact shape the whole recalc
 * story turns on. A fixture that cannot express the bug cannot catch it.
 *
 * Shared by the suites rather than copied into each: the second hand-rolled
 * OOXML skeleton in one directory is how two tests start disagreeing about what
 * a valid file looks like.
 */

import { zipSync } from 'fflate';

export interface MinimalWorkbook {
  /** Raw `<row>…</row>` markup for the single sheet. */
  rows: string;
  /** Sheet name, `S` by default. */
  name?: string;
  /** Adds a hyperlink on `A1` pointing here, with the rels part to match. */
  hyperlink?: string;
}

const enc = (s: string): Uint8Array => new TextEncoder().encode(s);

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function minimalXlsx(opts: MinimalWorkbook): ArrayBuffer {
  const name = opts.name ?? 'S';
  const links = opts.hyperlink ? `<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>` : '';

  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': enc(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ),
    '_rels/.rels': enc(
      `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rIdW" Type="${DOC}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    ),
    'xl/workbook.xml': enc(
      `<?xml version="1.0"?><workbook xmlns="${NS}" xmlns:r="${DOC}"><sheets><sheet name="${name}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': enc(
      `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${DOC}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
    'xl/worksheets/sheet1.xml': enc(
      `<?xml version="1.0"?><worksheet xmlns="${NS}" xmlns:r="${DOC}"><sheetData>${opts.rows}</sheetData>${links}</worksheet>`,
    ),
  };

  if (opts.hyperlink) {
    files['xl/worksheets/_rels/sheet1.xml.rels'] = enc(
      `<?xml version="1.0"?><Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${DOC}/hyperlink" Target="${opts.hyperlink}" TargetMode="External"/></Relationships>`,
    );
  }

  const z = zipSync(files);
  return z.buffer.slice(z.byteOffset, z.byteOffset + z.byteLength) as ArrayBuffer;
}
