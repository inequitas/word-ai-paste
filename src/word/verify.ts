import type { ReadBackParagraph } from './adapter';

/**
 * Phase 3 of the insert (see insert.ts): compare what Word actually has,
 * read back in document order, with what we meant to insert. Pure — no
 * Office.js — so it is unit-tested directly and shared by the self-test.
 */

/** One paragraph we meant to insert, in the order we meant it. */
export interface ExpectedParagraph {
  blockIndex: number;
  text: string;
  listItem: boolean;
  /** 1-based number of the list run this item belongs to (list items only). */
  run: number | null;
  level: number;
  ordered: boolean;
  /** Expected `styleBuiltIn`, or null when not checkable (custom body style). */
  style: string | null;
}

export type ListRunMode = 'list' | 'per-item';

/** What phase 2 did with one contiguous list run. */
export interface ListRunReport {
  run: number;
  listIndex: number;
  ordered: boolean;
  itemCount: number;
  /** Index into the expected-paragraph sequence of the run's first item. */
  firstExpected: number;
  /** "list": one Word list for the whole run; "per-item": from `fallbackFrom` on, every item is its own one-item list. */
  mode: ListRunMode;
  /** Item index (0-based, within the run) where the per-item fallback started; null in "list" mode. */
  fallbackFrom: number | null;
  listId: number | null;
}

const MAX_MISMATCHES = 25;

/** Collapses Word's line-break / whitespace variants so text compares reliably. */
export function normalizeText(text: string): string {
  return text
    .replace(/[\r\n]+/g, '\n')
    .replace(/[ \t ]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .trim();
}

export function snippet(text: string, max = 40): string {
  const t = normalizeText(text).replace(/\n/g, ' ');
  return `"${t.length > max ? `${t.slice(0, max)}…` : t}"`;
}

/** Text order: null when every paragraph matches, otherwise a description of the first difference. */
export function firstTextDifference(expected: ExpectedParagraph[], actual: ReadBackParagraph[]): string | null {
  const n = Math.min(expected.length, actual.length);
  for (let i = 0; i < n; i++) {
    if (normalizeText(expected[i].text) !== normalizeText(actual[i].text)) {
      return `paragraph #${i + 1}: expected ${snippet(expected[i].text)}, found ${snippet(actual[i].text)}`;
    }
  }
  if (expected.length !== actual.length) {
    return `expected ${expected.length} paragraphs, found ${actual.length}`;
  }
  return null;
}

/** The list-item paragraphs (read back) of one run, by position. */
export function runItems(run: ListRunReport, actual: ReadBackParagraph[]): (ReadBackParagraph | undefined)[] {
  return Array.from({ length: run.itemCount }, (_, k) => actual[run.firstExpected + k]);
}

/** Distinct list ids of a run's items (null when an id could not be read). */
export function runListIds(run: ListRunReport, actual: ReadBackParagraph[]): (number | null)[] {
  return runItems(run, actual).map((p) => (p ? p.listId : null));
}

/**
 * Compares the read-back with the expected sequence: text order, list
 * membership, style, and list-id sharing within / between runs. Returns
 * human-readable mismatch lines (empty when everything matches).
 */
export function compareReadBack(expected: ExpectedParagraph[], actual: ReadBackParagraph[], runs: ListRunReport[]): string[] {
  const out: string[] = [];
  const push = (line: string): void => {
    out.push(line);
  };

  const orderProblem = firstTextDifference(expected, actual);
  if (orderProblem) push(`text order: ${orderProblem}`);

  const n = Math.min(expected.length, actual.length);
  for (let i = 0; i < n; i++) {
    const e = expected[i];
    const a = actual[i];
    const label = `paragraph #${i + 1} ${snippet(e.text)}`;
    if (e.listItem && !a.isListItem) push(`${label}: should be an item of list ${e.run}, but is not a list item`);
    if (!e.listItem && a.isListItem) push(`${label}: should not be a list item, but is one`);
    if (e.style && a.styleBuiltIn !== e.style) push(`${label}: style should be ${e.style}, found ${a.styleBuiltIn}`);
    if (e.listItem && a.isListItem && a.level !== null && a.level !== e.level) {
      push(`${label}: list level should be ${e.level}, found ${a.level}`);
    }
  }

  const idsKnown = actual.some((p) => p.listId !== null);
  if (idsKnown) {
    let previous: { run: number; id: number } | null = null;
    for (const run of runs) {
      const ids = runListIds(run, actual);
      const distinct = Array.from(new Set(ids.filter((id): id is number => id !== null)));
      if (run.mode === 'list' && (distinct.length > 1 || ids.some((id) => id === null))) {
        push(`list ${run.run}: its ${run.itemCount} items do not share one list id (ids: ${ids.map((id) => id ?? '-').join(', ')})`);
      }
      const firstId = ids[0];
      if (previous && firstId !== null && firstId !== undefined && firstId === previous.id) {
        push(`list ${run.run} continues list ${previous.run} (same list id ${firstId}) instead of starting a new list`);
      }
      const lastId = ids[ids.length - 1];
      previous = lastId !== null && lastId !== undefined ? { run: run.run, id: lastId } : null;
    }
  }

  if (out.length > MAX_MISMATCHES) {
    const extra = out.length - MAX_MISMATCHES;
    return [...out.slice(0, MAX_MISMATCHES), `…and ${extra} more`];
  }
  return out;
}

/**
 * Indents of Word's own bullet button: level n puts the text at 36 + 36·n pt
 * and the bullet at 18 + 36·n pt (0.5" / 0.25" steps).
 */
export function listLevelIndents(level: number): { textIndent: number; bulletIndent: number } {
  return { textIndent: 36 + 36 * level, bulletIndent: 18 + 36 * level };
}

/**
 * Arguments for `list.setLevelIndents(level, textIndent, bulletIndentRelative)`.
 * Word treats the third argument as relative to the text indent (it becomes
 * the paragraph's `firstLineIndent`), as verified in Word for Mac 16.113.
 */
export function listLevelIndentArgs(level: number): { textIndent: number; bulletIndentRelative: number } {
  return { textIndent: 36 + 36 * level, bulletIndentRelative: -18 };
}

const INDENT_TOLERANCE_PT = 1.5;

/**
 * Does a read-back list item sit where Word's bullet button would put it?
 * `leftIndent` is the text position; `leftIndent + firstLineIndent` is the
 * bullet position (a hanging indent has a negative firstLineIndent).
 * Returns null when it matches, otherwise a description.
 */
export function indentProblem(p: ReadBackParagraph, level: number): string | null {
  const want = listLevelIndents(level);
  if (p.leftIndent === null || p.firstLineIndent === null) return 'indents not readable';
  const bullet = p.leftIndent + p.firstLineIndent;
  const ok =
    Math.abs(p.leftIndent - want.textIndent) <= INDENT_TOLERANCE_PT && Math.abs(bullet - want.bulletIndent) <= INDENT_TOLERANCE_PT;
  if (ok) return null;
  return (
    `level ${level}: leftIndent=${round(p.leftIndent)}, firstLineIndent=${round(p.firstLineIndent)} ` +
    `(text at ${round(p.leftIndent)}pt, bullet at ${round(bullet)}pt; want ${want.textIndent}/${want.bulletIndent})`
  );
}

export function round(n: number): number {
  return Math.round(n * 10) / 10;
}
