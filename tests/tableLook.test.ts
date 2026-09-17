import { describe, it, expect } from 'vitest';
import { formatLook, lookMismatches, parseTblLook, type TableLook } from '../src/word/tableLook';

describe('parseTblLook', () => {
  it('parses explicit attributes', () => {
    const xml = '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true
    });
  });

  it('parses hex bitmask 04A0 (firstRow, firstColumn, noVBand)', () => {
    const xml = '<w:tblLook w:val="04A0"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true
    });
  });

  it('parses hex bitmask 0520 (firstRow + lastColumn + noVBand)', () => {
    // 0x0520 = 0x0020 (firstRow) + 0x0100 (lastColumn) + 0x0400 (noVBand)
    const xml = '<w:tblLook w:val="0520"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: false,
      lastColumn: true,
      noHBand: false,
      noVBand: true
    });
  });

  it('parses case-insensitive attribute values (true/false)', () => {
    const xml = '<w:tblLook w:firstRow="true" w:lastRow="false" w:firstColumn="true" w:lastColumn="false" w:noHBand="false" w:noVBand="true"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true
    });
  });

  it('parses on/off attribute values', () => {
    const xml = '<w:tblLook w:firstRow="on" w:lastRow="off" w:firstColumn="on" w:lastColumn="off" w:noHBand="off" w:noVBand="on"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true
    });
  });

  it('returns null when no tblLook element exists', () => {
    const xml = '<w:tbl><w:tblPr></w:tblPr></w:tbl>';
    const look = parseTblLook(xml);
    expect(look).toBeNull();
  });

  it('handles tblLook with only hex val (no explicit attributes)', () => {
    const xml = '<w:tblLook w:val="0000"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: false,
      lastRow: false,
      firstColumn: false,
      lastColumn: false,
      noHBand: false,
      noVBand: false
    });
  });

  it('handles uppercase hex values', () => {
    const xml = '<w:tblLook w:val="04A0"/>';
    const look = parseTblLook(xml);
    expect(look).toEqual({
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true
    });
  });
});

describe('lookMismatches', () => {
  const base: TableLook = {
    firstRow: true,
    lastRow: false,
    firstColumn: true,
    lastColumn: false,
    noHBand: false,
    noVBand: false
  };

  it('returns empty array when tables match', () => {
    const mismatches = lookMismatches(base, base);
    expect(mismatches).toEqual([]);
  });

  it('detects single field mismatch', () => {
    const actual = { ...base, firstColumn: false };
    const mismatches = lookMismatches(actual, base);
    expect(mismatches).toEqual(['firstColumn']);
  });

  it('detects multiple field mismatches', () => {
    const actual = { ...base, firstColumn: false, noVBand: true };
    const mismatches = lookMismatches(actual, base);
    expect(mismatches).toContain('firstColumn');
    expect(mismatches).toContain('noVBand');
  });
});

describe('formatLook', () => {
  it('formats a table look as a compact string', () => {
    const look: TableLook = {
      firstRow: true,
      lastRow: false,
      firstColumn: true,
      lastColumn: false,
      noHBand: false,
      noVBand: true
    };
    const formatted = formatLook(look);
    expect(formatted).toBe('firstRow=1 lastRow=0 firstColumn=1 lastColumn=0 noHBand=0 noVBand=1');
  });

  it('formats all false', () => {
    const look: TableLook = {
      firstRow: false,
      lastRow: false,
      firstColumn: false,
      lastColumn: false,
      noHBand: false,
      noVBand: false
    };
    const formatted = formatLook(look);
    expect(formatted).toBe('firstRow=0 lastRow=0 firstColumn=0 lastColumn=0 noHBand=0 noVBand=0');
  });

  it('formats all true', () => {
    const look: TableLook = {
      firstRow: true,
      lastRow: true,
      firstColumn: true,
      lastColumn: true,
      noHBand: true,
      noVBand: true
    };
    const formatted = formatLook(look);
    expect(formatted).toBe('firstRow=1 lastRow=1 firstColumn=1 lastColumn=1 noHBand=1 noVBand=1');
  });
});
