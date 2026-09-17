import { describe, it, expect } from 'vitest';
import type { ReadBackParagraph } from '../src/word/adapter';
import {
  compareReadBack,
  firstTextDifference,
  indentProblem,
  listLevelIndents,
  normalizeText,
  type ExpectedParagraph,
  type ListRunReport
} from '../src/word/verify';

const exp = (text: string, extra: Partial<ExpectedParagraph> = {}): ExpectedParagraph => ({
  blockIndex: 0,
  text,
  listItem: false,
  run: null,
  level: 0,
  ordered: false,
  style: 'Normal',
  ...extra
});
const item = (text: string, run: number, level = 0): ExpectedParagraph =>
  exp(text, { listItem: true, run, level, style: 'ListParagraph' });

const got = (text: string, extra: Partial<ReadBackParagraph> = {}): ReadBackParagraph => ({
  text,
  isListItem: false,
  styleBuiltIn: 'Normal',
  leftIndent: 0,
  firstLineIndent: 0,
  listId: null,
  level: null,
  listString: null,
  ...extra
});
const gotItem = (text: string, listId: number, level = 0): ReadBackParagraph =>
  got(text, { isListItem: true, styleBuiltIn: 'ListParagraph', listId, level, leftIndent: 36, firstLineIndent: -18, listString: '•' });

const run = (n: number, firstExpected: number, itemCount: number, extra: Partial<ListRunReport> = {}): ListRunReport => ({
  run: n,
  listIndex: n,
  ordered: false,
  itemCount,
  firstExpected,
  mode: 'list',
  fallbackFrom: null,
  listId: null,
  ...extra
});

describe('normalizeText', () => {
  it("treats Word's vertical-tab line break, CR and non-breaking spaces like their plain forms", () => {
    expect(normalizeText('ab')).toBe('a\nb');
    expect(normalizeText('  a  b\r')).toBe('a b');
    expect(normalizeText('line one \n line two')).toBe('line one\nline two');
  });
});

describe('compareReadBack', () => {
  const expected = [exp('Intro'), item('a', 1), item('b', 1), exp('')];

  it('finds nothing when the read-back matches', () => {
    const actual = [got('Intro'), gotItem('a', 7), gotItem('b', 7), got('')];
    expect(compareReadBack(expected, actual, [run(1, 1, 2)])).toEqual([]);
    expect(firstTextDifference(expected, actual)).toBeNull();
  });

  it('reports the first text-order difference with its position', () => {
    const actual = [got('Intro'), gotItem('a', 7), got(''), got('b')];
    const m = compareReadBack(expected, actual, [run(1, 1, 2)]);
    expect(m[0]).toBe('text order: paragraph #3: expected "b", found ""');
  });

  it('reports a paragraph count difference', () => {
    expect(firstTextDifference(expected, [got('Intro')])).toBe('expected 4 paragraphs, found 1');
  });

  it('reports list membership in both directions', () => {
    const actual = [got('Intro', { isListItem: true, listId: 3 }), gotItem('a', 7), got('b', { styleBuiltIn: 'ListParagraph' }), got('')];
    const m = compareReadBack(expected, actual, [run(1, 1, 2)]);
    expect(m).toContain('paragraph #1 "Intro": should not be a list item, but is one');
    expect(m).toContain('paragraph #3 "b": should be an item of list 1, but is not a list item');
  });

  it('reports style and level differences', () => {
    const actual = [got('Intro', { styleBuiltIn: 'Heading1' }), gotItem('a', 7), gotItem('b', 7, 2), got('')];
    const m = compareReadBack([exp('Intro'), item('a', 1), item('b', 1, 1), exp('')], actual, [run(1, 1, 2)]);
    expect(m).toContain('paragraph #1 "Intro": style should be Normal, found Heading1');
    expect(m).toContain('paragraph #3 "b": list level should be 1, found 2');
  });

  it('does not compare the style of a custom body style', () => {
    const m = compareReadBack([exp('Intro', { style: null })], [got('Intro', { styleBuiltIn: 'Other' })], []);
    expect(m).toEqual([]);
  });

  it('reports a run whose items are on different lists — unless the per-item fallback was used', () => {
    const actual = [got('Intro'), gotItem('a', 7), gotItem('b', 8), got('')];
    expect(compareReadBack(expected, actual, [run(1, 1, 2)])).toContain('list 1: its 2 items do not share one list id (ids: 7, 8)');
    expect(compareReadBack(expected, actual, [run(1, 1, 2, { mode: 'per-item', fallbackFrom: 1 })])).toEqual([]);
  });

  it('reports a second list that continues the first one', () => {
    const e = [item('a', 1), item('b', 2)];
    const actual = [gotItem('a', 7), gotItem('b', 7)];
    expect(compareReadBack(e, actual, [run(1, 0, 1), run(2, 1, 1)])).toEqual([
      'list 2 continues list 1 (same list id 7) instead of starting a new list'
    ]);
  });

  it('skips list-id checks when no ids could be read', () => {
    const actual = [got('Intro'), gotItem('a', 7), gotItem('b', 7), got('')].map((p) => ({ ...p, listId: null }));
    expect(compareReadBack(expected, actual, [run(1, 1, 2)])).toEqual([]);
  });

  it('caps the number of reported mismatches', () => {
    const many = Array.from({ length: 40 }, (_, i) => exp(`p${i}`));
    const actual = many.map((e) => got(e.text, { isListItem: true }));
    const m = compareReadBack(many, actual, []);
    expect(m).toHaveLength(26);
    expect(m[25]).toBe('…and 15 more');
  });
});

describe('list indents', () => {
  it("uses Word's bullet-button geometry", () => {
    expect(listLevelIndents(0)).toEqual({ textIndent: 36, bulletIndent: 18 });
    expect(listLevelIndents(2)).toEqual({ textIndent: 108, bulletIndent: 90 });
  });

  it('accepts a hanging indent that puts the text at 36pt and the bullet at 18pt', () => {
    expect(indentProblem(gotItem('a', 1), 0)).toBeNull();
    expect(indentProblem(got('a', { leftIndent: 72.4, firstLineIndent: -18.2 }), 1)).toBeNull();
  });

  it('describes an indent that is too deep', () => {
    expect(indentProblem(got('a', { leftIndent: 90, firstLineIndent: -18 }), 0)).toBe(
      'level 0: leftIndent=90, firstLineIndent=-18 (text at 90pt, bullet at 72pt; want 36/18)'
    );
    expect(indentProblem(got('a', { leftIndent: 36, firstLineIndent: 18 }), 0)).toContain('bullet at 54pt');
    expect(indentProblem(got('a', { leftIndent: null }), 0)).toBe('indents not readable');
  });
});
