import type { Block, Inline, ListItemBlock, TableBlock } from '../model';
import type {
  ParagraphIntent,
  ReadBackResult,
  TableStyleSettings,
  WordDocument,
  WordListHandle,
  WordParagraphHandle,
  WordTableHandle
} from './adapter';
import {
  compareReadBack,
  listLevelIndents,
  type ExpectedParagraph,
  type ListRunReport
} from './verify';

export type BodyStyleOption = { kind: 'builtin' } | { kind: 'custom'; name: string };

export interface InsertOptions {
  bodyStyle: BodyStyleOption;
  tableStyle: TableStyleSettings;
}

export const DEFAULT_TABLE_STYLE: TableStyleSettings = {
  // Kevin's house style: Word's own "Grid Table 4" gallery swatch with its
  // default toggles (header row + first column emphasised, banded rows).
  styleBuiltIn: 'GridTable4',
  customName: null,
  headerRowCount: 1,
  styleFirstColumn: true,
  styleBandedRows: true,
  styleBandedColumns: false,
  styleLastColumn: false,
  styleTotalRow: false
};

export const DEFAULT_INSERT_OPTIONS: InsertOptions = {
  bodyStyle: { kind: 'builtin' },
  tableStyle: DEFAULT_TABLE_STYLE
};

const MONOSPACE_FONT = 'Consolas';
const MAX_HEADING_LEVEL = 9;

// ---------------------------------------------------------------------------
// Diagnostics: step log + a rich error the task pane can display and copy
// ---------------------------------------------------------------------------

export interface StepLogEntry {
  blockIndex: number;
  blockType: string;
  action: string;
}

export function formatStep(s: StepLogEntry): string {
  return `block ${s.blockIndex} ${s.blockType}: ${s.action}`;
}

export interface OfficeErrorInfo {
  code?: string;
  message: string;
  errorLocation?: string;
  statement?: string;
}

/** Best-effort extraction from an Office.js `OfficeExtension.Error` (or anything else that gets thrown). */
export function extractOfficeErrorInfo(err: unknown): OfficeErrorInfo {
  if (err && typeof err === 'object') {
    const e = err as { code?: string; message?: string; debugInfo?: { errorLocation?: string; statement?: string } };
    return {
      code: e.code,
      message: e.message || String(err),
      errorLocation: e.debugInfo?.errorLocation,
      statement: e.debugInfo?.statement
    };
  }
  return { message: String(err) };
}

/** One-line summary of an Office.js error: "code message @ errorLocation". */
export function describeError(err: unknown): string {
  const info = extractOfficeErrorInfo(err);
  return `${info.code && info.code !== info.message ? `${info.code} ` : ''}${info.message}${
    info.errorLocation ? ` @ ${info.errorLocation}` : ''
  }`;
}

/** Thrown by insertBlocks when phase 1 (the text) fails. Carries the full step log and the step that was in progress. */
export class InsertError extends Error {
  readonly steps: StepLogEntry[];
  readonly failedStep: StepLogEntry | null;
  readonly officeError: OfficeErrorInfo;
  readonly rawError: unknown;

  constructor(steps: StepLogEntry[], failedStep: StepLogEntry | null, rawError: unknown) {
    const info = extractOfficeErrorInfo(rawError);
    const where = failedStep ? `block ${failedStep.blockIndex} (${failedStep.blockType}): ${failedStep.action}` : 'setup';
    super(`Insert failed at ${where}: ${info.message}`);
    this.name = 'InsertError';
    this.steps = steps;
    this.failedStep = failedStep;
    this.officeError = info;
    this.rawError = rawError;
  }
}

/** What insertBlocks did, and what the read-back afterwards found. */
export interface InsertReport {
  steps: StepLogEntry[];
  /** Problems found by the read-back (phase 3); empty when everything matched. */
  mismatches: string[];
  /** The paragraphs we meant to insert, in order. */
  expected: ExpectedParagraph[];
  /** What phase 2 did with each list run. */
  runs: ListRunReport[];
  /** The read-back, or null if it was skipped or failed (see `verifyError`). */
  readBack: ReadBackResult | null;
  verifyError: string | null;
  /** Paragraphs the safety net had to detach from a list. */
  repaired: number;
}

/**
 * Two-phase insert
 * ----------------
 * Real Word for Mac (16.113) does not add a paragraph to a list when you
 * call `list.insertParagraph(text, "End")`: the paragraph lands directly
 * after the list's current range (still just item 1), so a 6-item list
 * came out as item 1, [everything after the list], item 6, 5, 4, 3, 2 —
 * and `"After"` on a list inserts at that same spot. So lists are no
 * longer built while inserting:
 *
 * Phase 1 — all text, in order. Every block, list items included, becomes
 * an ordinary paragraph created with `paragraph.insertParagraph(text,
 * "After")`, chained from the previous one (proven to keep order and
 * styles on Mac), synced per block. List items get the ListParagraph
 * style; we remember each one's run (contiguous items with the same
 * listIndex) and level. No list exists yet, so nothing can inherit list
 * membership by accident.
 *
 * Phase 2 — turn each run into a list: `startNewList()` on its first item,
 * read the new list's id, `attachToList(id, level)` for every later item
 * (sync + read back after each), then configure every level used (bullet
 * or "1." numbering, and the indents of Word's own bullet button), each in
 * its own try + sync. If `attachToList` throws or leaves the paragraph
 * outside the list, the remaining items each get their own one-item list
 * (`startNewList()` again), with `setLevelStartingNumber` keeping ordered
 * numbers reading 1, 2, 3 ("per-item-lists fallback"). Phase 2 never
 * throws: the text is already in place, and phase 3 reports what's wrong.
 *
 * Phase 3 — verify. The safety net detaches any non-list paragraph that
 * ended up a list item anyway, then the whole inserted range is read back
 * in document order and compared with what we meant (text order, list
 * membership, style, one list id per run). Mismatches go to the step log,
 * console.warn and `InsertReport.mismatches`.
 */
export async function insertBlocks(
  doc: WordDocument,
  blocks: Block[],
  options: InsertOptions = DEFAULT_INSERT_OPTIONS
): Promise<InsertReport> {
  const steps: StepLogEntry[] = [];
  const report: InsertReport = {
    steps,
    mismatches: [],
    expected: [],
    runs: [],
    readBack: null,
    verifyError: null,
    repaired: 0
  };
  if (!blocks.length) return report;

  const log = (blockIndex: number, blockType: string, action: string): void => {
    steps.push({ blockIndex, blockType, action });
  };

  const phase1 = await insertText(doc, blocks, options, log, steps, report.expected);
  report.runs = await convertListRuns(doc, phase1.runs, log);

  try {
    log(-1, 'verify', 'verifyAndRepairListItems');
    report.repaired = await doc.verifyAndRepairListItems(phase1.intents);
    if (report.repaired) log(-1, 'verify', `repair: detached ${report.repaired} paragraph(s) that should not be list items`);
  } catch (err) {
    log(-1, 'verify', `repair pass failed (${describeError(err)})`);
  }

  if (phase1.handles.length) {
    try {
      log(-1, 'verify', 'readBack');
      const readBack = await doc.readBack(phase1.handles[0], phase1.handles[phase1.handles.length - 1]);
      report.readBack = readBack;
      for (const note of readBack.notes) log(-1, 'verify', `note: ${note}`);
      report.mismatches = compareReadBack(report.expected, readBack.paragraphs, report.runs);
    } catch (err) {
      report.verifyError = describeError(err);
      log(-1, 'verify', `readBack failed (${report.verifyError})`);
    }
  }

  for (const m of report.mismatches) {
    log(-1, 'verify', `mismatch: ${m}`);
    console.warn(`AI Paste: ${m}`);
  }
  log(-1, 'verify', report.mismatches.length ? `${report.mismatches.length} mismatch(es)` : 'read-back matches');
  return report;
}

// ---------------------------------------------------------------------------
// Phase 1: all text, in order
// ---------------------------------------------------------------------------

interface PendingItem {
  handle: WordParagraphHandle;
  block: ListItemBlock;
  blockIndex: number;
}

interface PendingRun {
  run: number;
  listIndex: number;
  firstExpected: number;
  items: PendingItem[];
}

interface Phase1Result {
  /** Every paragraph we inserted (or reused), in document order. */
  handles: WordParagraphHandle[];
  intents: ParagraphIntent[];
  runs: PendingRun[];
}

type Log = (blockIndex: number, blockType: string, action: string) => void;

/** Where the next paragraph goes. */
type Anchor =
  | { kind: 'paragraph'; paragraph: WordParagraphHandle; reusable: boolean }
  | { kind: 'table'; table: WordTableHandle };

async function insertText(
  doc: WordDocument,
  blocks: Block[],
  options: InsertOptions,
  log: Log,
  steps: StepLogEntry[],
  expected: ExpectedParagraph[]
): Promise<Phase1Result> {
  const result: Phase1Result = { handles: [], intents: [], runs: [] };
  const fail = (err: unknown): never => {
    throw new InsertError(steps, steps[steps.length - 1] ?? null, err);
  };

  let cursorIsListItem = false;
  let anchor: Anchor;
  try {
    log(-1, 'setup', 'getCursor');
    const cursor = await doc.getCursor();
    anchor = { kind: 'paragraph', paragraph: cursor.paragraph, reusable: cursor.isEmpty };
    cursorIsListItem = cursor.isListItem;
    if (cursorIsListItem) log(-1, 'setup', 'cursor paragraph is a list item');
  } catch (err) {
    return fail(err);
  }

  /** Paragraphs chained directly from a list-item cursor may inherit its list; checked once, right away. */
  let checkInheritedList = cursorIsListItem;

  /** Gets (or creates) the paragraph for the next piece of content, with `text` set from the start. */
  const nextParagraph = async (text: string, blockIndex: number, blockType: string): Promise<WordParagraphHandle> => {
    let p: WordParagraphHandle;
    if (anchor.kind === 'paragraph' && anchor.reusable) {
      p = anchor.paragraph;
      if (checkInheritedList) {
        // The empty cursor paragraph is itself a list item (confirmed by getCursor's load).
        log(blockIndex, blockType, 'detachFromList (reused cursor paragraph)');
        p.detachFromList();
        checkInheritedList = false;
      }
      if (text) p.insertRun(text, { bold: false, italic: false });
    } else if (anchor.kind === 'table') {
      p = anchor.table.insertParagraphAfter(text);
    } else {
      p = anchor.paragraph.insertParagraphAfter(text);
      if (checkInheritedList) {
        checkInheritedList = false;
        await doc.commit();
        const [state] = await doc.readListState([p]);
        if (state.isListItem) {
          log(blockIndex, blockType, 'detachFromList (inherited from the list-item cursor paragraph)');
          p.detachFromList();
          await doc.commit();
        }
      }
    }
    anchor = { kind: 'paragraph', paragraph: p, reusable: false };
    result.handles.push(p);
    return p;
  };

  const addExpected = (p: WordParagraphHandle, e: ExpectedParagraph, reapplyStyle: (() => void) | null): void => {
    expected.push(e);
    if (reapplyStyle) {
      reapplyStyle();
      result.intents.push({ handle: p, intendedListItem: false, reapplyStyle });
    } else {
      result.intents.push({ handle: p, intendedListItem: true });
    }
  };

  let currentRun: PendingRun | null = null;

  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index];
    try {
      if (block.type !== 'listItem') currentRun = null;

      if (block.type === 'table') {
        let tableAnchor: WordParagraphHandle;
        if (anchor.kind === 'table') {
          // Two tables in a row: Word needs a paragraph between them.
          log(index, 'table', 'insertParagraphAfter(table) spacer');
          tableAnchor = await nextParagraph('', index, 'table');
          addExpected(tableAnchor, bodyExpectation(index, '', options), () => applyBodyStyle(tableAnchor, options));
        } else {
          tableAnchor = anchor.paragraph;
        }
        log(index, 'table', 'insertTable');
        const table = tableAnchor.insertTableAfter(tableValues(block));
        await doc.commit();

        const headerRowCount = block.header ? options.tableStyle.headerRowCount : 0;
        log(index, 'table', 'applyStyle');
        await table.applyStyle({ ...options.tableStyle, headerRowCount });
        // The next block (if any) is created directly after the table, with
        // its own text — no empty paragraph first (normalize.ts already adds
        // the house-style blank line after a table).
        anchor = { kind: 'table', table };
        checkInheritedList = false;
        continue;
      }

      if (block.type === 'listItem') {
        if (!currentRun || currentRun.listIndex !== block.listIndex) {
          currentRun = { run: result.runs.length + 1, listIndex: block.listIndex, firstExpected: expected.length, items: [] };
          result.runs.push(currentRun);
        }
        const first = block.inlines[0];
        log(index, 'listItem', `insertParagraphAfter (list ${currentRun.run} item ${currentRun.items.length + 1}, level ${block.level})`);
        const p = await nextParagraph(first ? first.text : '', index, 'listItem');
        if (first) p.setWholeTextFormat(runFormat(first));
        appendRemainingInlines(p, block.inlines);
        p.setStyleBuiltIn('ListParagraph');
        addExpected(
          p,
          {
            blockIndex: index,
            text: blockText(block.inlines),
            listItem: true,
            run: currentRun.run,
            level: block.level,
            ordered: block.ordered,
            style: 'ListParagraph'
          },
          null
        );
        currentRun.items.push({ handle: p, block, blockIndex: index });
        await doc.commit();
        continue;
      }

      if (block.type === 'code') {
        // One paragraph per line, not one paragraph with embedded line
        // breaks, so each line is independently selectable/editable.
        for (const line of block.text.split('\n')) {
          log(index, 'code', 'insertParagraphAfter(line)');
          const linePara = await nextParagraph(line, index, 'code');
          if (line) linePara.setWholeTextFormat({ bold: false, italic: false, fontName: MONOSPACE_FONT });
          addExpected(linePara, bodyExpectation(index, line, options), () => applyBodyStyle(linePara, options));
        }
        await doc.commit();
        continue;
      }

      // heading / paragraph (including blank spacers) / quote
      const first = block.inlines[0];
      log(index, block.type, 'insertParagraphAfter');
      const p = await nextParagraph(first ? first.text : '', index, block.type);
      if (first) p.setWholeTextFormat(runFormat(first));
      appendRemainingInlines(p, block.inlines);
      const text = blockText(block.inlines);

      switch (block.type) {
        case 'heading': {
          const style = headingStyleName(block.level);
          addExpected(p, { ...bodyExpectation(index, text, options), style }, () => p.setStyleBuiltIn(style));
          break;
        }
        case 'quote':
          addExpected(p, { ...bodyExpectation(index, text, options), style: 'Quote' }, () => p.setStyleBuiltIn('Quote'));
          break;
        case 'paragraph':
        default:
          addExpected(p, bodyExpectation(index, text, options), () => applyBodyStyle(p, options));
          break;
      }
      await doc.commit();
    } catch (err) {
      return fail(err);
    }
  }

  return result;
}

function bodyExpectation(blockIndex: number, text: string, options: InsertOptions): ExpectedParagraph {
  return {
    blockIndex,
    text,
    listItem: false,
    run: null,
    level: 0,
    ordered: false,
    style: options.bodyStyle.kind === 'builtin' ? 'Normal' : null
  };
}

// ---------------------------------------------------------------------------
// Phase 2: turn list runs into Word lists
// ---------------------------------------------------------------------------

async function convertListRuns(doc: WordDocument, runs: PendingRun[], log: Log): Promise<ListRunReport[]> {
  const reports: ListRunReport[] = [];
  for (const run of runs) {
    reports.push(await convertListRun(doc, run, log));
  }
  return reports;
}

async function convertListRun(doc: WordDocument, run: PendingRun, log: Log): Promise<ListRunReport> {
  const items = run.items;
  const ordered = items[0].block.ordered;
  const at = items[0].blockIndex;
  const tag = `list ${run.run}`;
  const note = (action: string, blockIndex = at): void => log(blockIndex, 'list', `${tag}: ${action}`);
  const report: ListRunReport = {
    run: run.run,
    listIndex: run.listIndex,
    ordered,
    itemCount: items.length,
    firstExpected: run.firstExpected,
    mode: 'list',
    fallbackFrom: null,
    listId: null
  };
  const ordinals = computeOrdinals(items.map((i) => i.block));

  let list: WordListHandle | null = null;
  let fallbackFrom: number | null = null;

  try {
    note('startNewList');
    list = items[0].handle.startNewList();
    await doc.commit();
    note('load list id');
    report.listId = await list.getId();
    note(`list id ${report.listId}`);
  } catch (err) {
    note(`startNewList / list id failed (${describeError(err)})`);
    list = null;
    fallbackFrom = 0;
  }

  if (list && report.listId !== null) {
    const listId = report.listId;
    if (items[0].block.level > 0) await setLevelSafely(doc, items[0].handle, items[0].block.level, (a) => note(a));

    for (let i = 1; i < items.length; i++) {
      const { handle, block, blockIndex } = items[i];
      try {
        note(`attachToList(${listId}, ${block.level}) item ${i + 1}`, blockIndex);
        handle.attachToList(listId, block.level);
        await doc.commit();
        const [state] = await doc.readListState([handle]);
        if (!state.isListItem) {
          note(`item ${i + 1}: attachToList left isListItem false`, blockIndex);
          fallbackFrom = i;
          break;
        }
        if (state.listId !== null && state.listId !== listId) {
          note(`item ${i + 1}: joined list id ${state.listId}, not ${listId}`, blockIndex);
        }
        if (state.level !== null && state.level !== block.level) {
          note(`item ${i + 1}: level is ${state.level}, setting ${block.level}`, blockIndex);
          await setLevelSafely(doc, handle, block.level, (a) => note(a, blockIndex));
        }
      } catch (err) {
        note(`item ${i + 1}: attachToList threw (${describeError(err)})`, blockIndex);
        fallbackFrom = i;
        break;
      }
    }

    const attached = items.slice(0, fallbackFrom ?? items.length);
    const levels = Array.from(new Set(attached.map((i) => i.block.level))).sort((a, b) => a - b);
    for (const level of levels) {
      const firstAtLevel = attached.find((i) => i.block.level === level)!;
      const start = firstAtLevel.block.start;
      await configureLevel(doc, list, level, ordered, start !== undefined && start !== 1 ? start : null, (a) => note(a));
    }
  }

  if (fallbackFrom !== null) {
    report.mode = 'per-item';
    report.fallbackFrom = fallbackFrom;
    note(`per-item-lists fallback (items ${fallbackFrom + 1}-${items.length})`);
    for (let i = fallbackFrom; i < items.length; i++) {
      const { handle, block, blockIndex } = items[i];
      const itemNote = (a: string): void => note(`item ${i + 1}: ${a}`, blockIndex);
      try {
        itemNote('startNewList');
        const own = handle.startNewList();
        await doc.commit();
        if (block.level > 0) await setLevelSafely(doc, handle, block.level, itemNote);
        await configureLevel(doc, own, block.level, block.ordered, block.ordered ? ordinals[i] : null, itemNote);
      } catch (err) {
        itemNote(`startNewList failed (${describeError(err)})`);
      }
    }
  }

  note(report.mode === 'list' ? `done (one list, ${items.length} items)` : `done (per-item lists from item ${fallbackFrom! + 1})`);
  return report;
}

async function setLevelSafely(doc: WordDocument, p: WordParagraphHandle, level: number, note: (a: string) => void): Promise<void> {
  try {
    note(`listItem.level = ${level}`);
    p.setListLevel(level);
    await doc.commit();
  } catch (err) {
    note(`listItem.level = ${level} failed (${describeError(err)})`);
  }
}

/** Bullet/number format, indents and (optionally) starting number of one level — each its own try + sync. */
async function configureLevel(
  doc: WordDocument,
  list: WordListHandle,
  level: number,
  ordered: boolean,
  startingNumber: number | null,
  note: (a: string) => void
): Promise<void> {
  const step = async (label: string, run: () => void): Promise<void> => {
    try {
      note(label);
      run();
      await doc.commit();
    } catch (err) {
      note(`${label} failed (${describeError(err)})`);
      console.warn(`AI Paste: ${label} failed`, err);
    }
  };

  if (ordered) await step(`setLevelNumbering(${level}, arabic)`, () => list.setLevelNumbering(level));
  else await step(`setLevelBullet(${level}, solid)`, () => list.setLevelBullet(level));

  const { textIndent, bulletIndent } = listLevelIndents(level);
  await step(`setLevelIndents(${level}, ${textIndent}, ${bulletIndent})`, () => list.setLevelIndents(level, textIndent, bulletIndent));

  if (startingNumber !== null) {
    await step(`setLevelStartingNumber(${level}, ${startingNumber})`, () => list.setLevelStartingNumber(level, startingNumber));
  }
}

/**
 * The number each item would show in one proper list: counts per level,
 * restarting a level after any shallower item, honouring an explicit
 * `start` on the first item of a level.
 */
export function computeOrdinals(items: ListItemBlock[]): number[] {
  const counters: number[] = [];
  return items.map((item) => {
    counters.length = Math.min(counters.length, item.level + 1);
    const current = counters[item.level];
    counters[item.level] = current === undefined ? (item.start ?? 1) : current + 1;
    return counters[item.level];
  });
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function applyBodyStyle(p: WordParagraphHandle, options: InsertOptions): void {
  if (options.bodyStyle.kind === 'custom') p.setStyleByName(options.bodyStyle.name);
  else p.setStyleBuiltIn('Normal');
}

function headingStyleName(level: number): string {
  const clamped = Math.min(Math.max(Math.round(level), 1), MAX_HEADING_LEVEL);
  return `Heading${clamped}`;
}

function blockText(inlines: Inline[]): string {
  return inlines.map((i) => i.text).join('');
}

type RunFormatLike = { bold: boolean; italic: boolean; fontName?: string; link?: string };

function runFormat(inline: Inline): RunFormatLike {
  return {
    bold: !!inline.bold,
    italic: !!inline.italic,
    fontName: inline.code ? MONOSPACE_FONT : undefined,
    link: inline.link
  };
}

/** Appends every inline after the first (the first was already passed directly to the paragraph-creating call). */
function appendRemainingInlines(p: WordParagraphHandle, inlines: Inline[]): void {
  for (const inline of inlines.slice(1)) {
    if (!inline.text) continue;
    p.insertRun(inline.text, runFormat(inline));
  }
}

function tableValues(block: TableBlock): string[][] {
  const colCount = Math.max(
    block.header?.length ?? 0,
    block.rows.reduce((max, row) => Math.max(max, row.length), 0),
    1
  );
  const padRow = (row: Inline[][]): string[] => {
    const cells = row.map(cellText);
    while (cells.length < colCount) cells.push('');
    return cells.slice(0, colCount);
  };

  const rows: string[][] = [];
  if (block.header) rows.push(padRow(block.header));
  for (const row of block.rows) rows.push(padRow(row));
  if (!rows.length) rows.push(new Array(colCount).fill(''));
  return rows;
}

function cellText(cell: Inline[]): string {
  // Plain text per v1 (no per-cell rich runs) — see README limitations.
  return cell
    .map((i) => i.text)
    .join('')
    .replace(/\n/g, ' ')
    .trim();
}
