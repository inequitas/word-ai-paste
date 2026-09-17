/**
 * Parsing and formatting of Word's tblLook (table look) element.
 * The tblLook element controls the visual table style flags: which rows/columns
 * get special formatting. This module handles both explicit attributes and
 * hex-encoded bitmasks.
 */

export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

/**
 * Parses the first `<w:tblLook />` element found in the OOXML string.
 *
 * If explicit attributes are present (w:firstRow, w:lastRow, etc.), they are used.
 * Accepted values: "0"/"1", "true"/"false", "on"/"off".
 *
 * Otherwise, decodes the hex `w:val` bitmask:
 * - 0x0020: firstRow
 * - 0x0040: lastRow
 * - 0x0080: firstColumn
 * - 0x0100: lastColumn
 * - 0x0200: noHBand (no horizontal band)
 * - 0x0400: noVBand (no vertical band)
 *
 * Returns null if no tblLook element is found or parsing fails.
 */
export function parseTblLook(ooxml: string): TableLook | null {
  const match = /<w:tblLook\s+([^>]*)\/?>/i.exec(ooxml);
  if (!match) return null;

  const attrs = match[1];

  // Try to parse explicit attributes first
  const firstRow = parseAttr(attrs, 'w:firstRow');
  if (firstRow !== null) {
    // Explicit attributes are present
    return {
      firstRow,
      lastRow: parseAttr(attrs, 'w:lastRow') ?? false,
      firstColumn: parseAttr(attrs, 'w:firstColumn') ?? false,
      lastColumn: parseAttr(attrs, 'w:lastColumn') ?? false,
      noHBand: parseAttr(attrs, 'w:noHBand') ?? false,
      noVBand: parseAttr(attrs, 'w:noVBand') ?? false
    };
  }

  // Try hex bitmask
  const hexMatch = /w:val="([0-9A-Fa-f]+)"/.exec(attrs);
  if (!hexMatch) return null;

  const hexVal = parseInt(hexMatch[1], 16);
  return {
    firstRow: !!(hexVal & 0x0020),
    lastRow: !!(hexVal & 0x0040),
    firstColumn: !!(hexVal & 0x0080),
    lastColumn: !!(hexVal & 0x0100),
    noHBand: !!(hexVal & 0x0200),
    noVBand: !!(hexVal & 0x0400)
  };
}

/**
 * Parses a single boolean attribute value. Accepts "0"/"1", "true"/"false", "on"/"off".
 * Returns the boolean value if found, or null if not present or unparseable.
 */
function parseAttr(attrs: string, name: string): boolean | null {
  const regex = new RegExp(`${name}="([^"]*)"`, 'i');
  const match = regex.exec(attrs);
  if (!match) return null;
  const value = match[1].toLowerCase();
  if (value === '1' || value === 'true' || value === 'on') return true;
  if (value === '0' || value === 'false' || value === 'off') return false;
  return null;
}

/**
 * Compares two TableLook objects and returns an array of field names that differ.
 */
export function lookMismatches(actual: TableLook, target: TableLook): string[] {
  const mismatches: string[] = [];
  const fields: (keyof TableLook)[] = ['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand'];
  for (const field of fields) {
    if (actual[field] !== target[field]) {
      mismatches.push(field);
    }
  }
  return mismatches;
}

/**
 * Formats a TableLook as a compact string like "firstRow=1 lastRow=0 firstColumn=1 lastColumn=0 noHBand=0 noVBand=1".
 */
export function formatLook(l: TableLook): string {
  return [
    `firstRow=${l.firstRow ? 1 : 0}`,
    `lastRow=${l.lastRow ? 1 : 0}`,
    `firstColumn=${l.firstColumn ? 1 : 0}`,
    `lastColumn=${l.lastColumn ? 1 : 0}`,
    `noHBand=${l.noHBand ? 1 : 0}`,
    `noVBand=${l.noVBand ? 1 : 0}`
  ].join(' ');
}
