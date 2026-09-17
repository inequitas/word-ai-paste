import type { Block, Inline, ListItemBlock, TableBlock } from '../model';
import type {
  OoxmlListItemInput,
  WordDocument,
  WordListHandle,
  WordParagraphHandle,
  TableStyleSettings,
  ParagraphIntent
} from './adapter';

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

/** Thrown by insertBlocks on failure. Carries the full step log and the step that was in progress, for on-screen diagnostics. */
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

/**
 * List-numbering-restart strategy
 * --------------------------------
 * Every distinct `listIndex` produced by the parsers gets its own
 * `Word.List` object, created with `paragraph.startNewList()` for that
 * list's first item; later items — nested sub-items included — join with
 * `list.insertParagraph(text, "End")`. This is deliberately different from
 * the more obvious-looking `paragraph.attachToList(listId, level)` route:
 * `attachToList` needs the numeric list id as a plain argument, and reading
 * `list.id` back off a freshly created `Word.List` proxy needs its own
 * `load()` + `context.sync()`. Holding on to the in-memory `Word.List`
 * object returned by `startNewList()` sidesteps that, and a brand new
 * `Word.List` restarts numbering at 1 by construction.
 *
 * Step-wise execution
 * --------------------
 * An earlier version queued an entire paste into one batch with a single
 * final sync. Against real Word (16.113 on Mac) that produced a bare
 * "GeneralException" with no useful detail, and Kevin's own test pinpointed
 * it to the very first list item. This version syncs after every block
 * (`WordDocument.commit()`) instead, changes two things that were
 * suspected of causing it — every paragraph is now created WITH its first
 * run's text already passed to the creating call (never created empty and
 * filled afterwards, and in particular `startNewList()` is always called
 * on a paragraph that already has text), and a nested list item's `level`
 * is set only after its own commit, using `paragraph.listItem` (the
 * throwing variant, meaningful only once list membership is confirmed) —
 * and adds a whole-list OOXML fallback (`processListRun` below) for when
 * the paragraph/list API throws anyway.
 *
 * Leaving a list cleanly
 * -----------------------
 * In real Word, `paragraph.insertParagraphAfter()` on a paragraph that is
 * currently a list item usually creates *another list item*, and setting
 * `styleBuiltIn` afterwards does not reliably remove that list membership.
 * So for every block that is not a list item, this module tracks whichever
 * `Word.List` the "current position" belongs to (`currentList`) and gets
 * its paragraph via `list.insertParagraphAfter()` instead — which inserts
 * after the whole list rather than as one more member of it. Tables are
 * the one exception: not being paragraphs, they never inherit numPr.
 *
 * As a safety net on top of all of the above, every non-list paragraph is
 * tracked and re-checked with `WordDocument.verifyAndRepairListItems` after
 * the main loop — see that method's doc comment.
 */

export async function insertBlocks(
  doc: WordDocument,
  blocks: Block[],
  options: InsertOptions = DEFAULT_INSERT_OPTIONS
): Promise<void> {
  if (!blocks.length) return;

  const steps: StepLogEntry[] = [];
  const log = (blockIndex: number, blockType: string, action: string): void => {
    steps.push({ blockIndex, blockType, action });
  };
  const fail = (err: unknown): never => {
    throw new InsertError(steps, steps[steps.length - 1] ?? null, err);
  };

  let cursorInfo;
  try {
    log(-1, 'setup', 'getCursor');
    cursorInfo = await doc.getCursor();
  } catch (err) {
    return fail(err);
  }

  let cursor = cursorInfo.paragraph;
  let anchorReusable = cursorInfo.isEmpty;
  let currentList: WordListHandle | null = cursorInfo.list;
  const tracked: ParagraphIntent[] = [];

  /** A fresh paragraph, with `text` already set, guaranteed not to be a member of any list. */
  const paragraphOutsideList = (text: string): WordParagraphHandle => {
    const p = currentList ? currentList.insertParagraphAfter(text) : cursor.insertParagraphAfter(text);
    currentList = null;
    return p;
  };

  /** Gets (or creates) the paragraph for the next non-list block, with `text` set from the start. */
  const nextParagraph = (text: string): WordParagraphHandle => {
    if (anchorReusable) {
      anchorReusable = false;
      if (currentList) {
        // The reused cursor paragraph was itself an (empty) list item —
        // detach is only ever called here because CursorInfo.list was
        // populated from a genuine load in getCursor, never speculatively.
        cursor.detachFromList();
        currentList = null;
      }
      if (text) cursor.insertRun(text, { bold: false, italic: false });
      return cursor;
    }
    return paragraphOutsideList(text);
  };

  const trackNonList = (p: WordParagraphHandle, reapplyStyle: () => void): void => {
    reapplyStyle();
    tracked.push({ handle: p, intendedListItem: false, reapplyStyle });
  };

  /**
   * Inserts one whole list (all its items, including nested ones — a
   * contiguous run sharing the same listIndex) via the primary
   * paragraph/list API, falling back to a raw-OOXML insert (see ooxml.ts)
   * for whatever didn't get created if any step throws. Updates the outer
   * `cursor`/`currentList`/`tracked` on completion; only throws if BOTH
   * the primary path and the fallback fail.
   */
  const processListRun = async (items: ListItemBlock[], runStartIndex: number): Promise<void> => {
    const itemParagraphs: WordParagraphHandle[] = [];
    let list: WordListHandle | null = null;
    let item0Paragraph: WordParagraphHandle | null = null;
    let primaryFailed: unknown = null;

    try {
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const first = item.inlines[0];
        const text = first ? first.text : '';
        let p: WordParagraphHandle;

        if (i === 0) {
          p = nextParagraph(text);
          item0Paragraph = p;
          if (first) p.setWholeTextFormat(runFormat(first));
          appendRemainingInlines(p, item.inlines);
          p.setStyleBuiltIn('ListParagraph');
          log(runStartIndex, 'listItem', 'startNewList');
          list = p.startNewList();
        } else {
          log(runStartIndex + i, 'listItem', 'insertItemAfter');
          p = list!.insertItemAfter(text);
          if (first) p.setWholeTextFormat(runFormat(first));
          appendRemainingInlines(p, item.inlines);
          p.setStyleBuiltIn('ListParagraph');
        }

        await doc.commit();
        itemParagraphs.push(p);

        log(runStartIndex + i, 'listItem', item.ordered ? 'ensureNumberLevel' : 'ensureBulletLevel');
        if (item.ordered) await list!.ensureNumberLevel(item.level, item.start);
        else await list!.ensureBulletLevel(item.level);

        if (item.level > 0) {
          log(runStartIndex + i, 'listItem', `setListLevel(${item.level})`);
          p.setListLevel(item.level);
          await doc.commit();
        }
      }
    } catch (err) {
      primaryFailed = err;
    }

    if (!primaryFailed) {
      currentList = list;
      cursor = itemParagraphs[itemParagraphs.length - 1];
      for (const p of itemParagraphs) tracked.push({ handle: p, intendedListItem: true });
      log(runStartIndex, 'listItem', `list ${items[0].listIndex}: api`);
      return;
    }

    const errInfo = extractOfficeErrorInfo(primaryFailed);
    log(
      runStartIndex,
      'listItem',
      `list ${items[0].listIndex}: ooxml-fallback (${errInfo.message}${errInfo.errorLocation ? ` @ ${errInfo.errorLocation}` : ''})`
    );

    const alreadyDone = itemParagraphs.length;
    const remaining = items.slice(alreadyDone);
    const fallbackAnchor = alreadyDone === 0 ? item0Paragraph! : itemParagraphs[itemParagraphs.length - 1];
    const fallbackAnchorIsEmpty = alreadyDone === 0;
    const ooxmlInput: OoxmlListItemInput[] = remaining.map((b) => ({ level: b.level, ordered: b.ordered, inlines: b.inlines }));

    const result = await doc.insertOoxmlList(fallbackAnchor, fallbackAnchorIsEmpty, ooxmlInput);

    const allItemParagraphs = [...itemParagraphs, ...result.items];
    for (const p of allItemParagraphs) tracked.push({ handle: p, intendedListItem: true });

    currentList = result.list;
    cursor = allItemParagraphs[allItemParagraphs.length - 1] ?? cursor;
    log(runStartIndex, 'listItem', `list ${items[0].listIndex}: ooxml-fallback inserted ${remaining.length} item(s)`);
  };

  for (let index = 0; index < blocks.length; ) {
    const block = blocks[index];
    try {
      if (block.type === 'table') {
        const tableAnchor = cursor;
        anchorReusable = false;
        currentList = null;
        log(index, 'table', 'insertTable');
        const table = tableAnchor.insertTableAfter(tableValues(block));
        await doc.commit();

        const headerRowCount = block.header ? options.tableStyle.headerRowCount : 0;
        log(index, 'table', 'applyStyle');
        await table.applyStyle({ ...options.tableStyle, headerRowCount });

        const hasMore = index < blocks.length - 1;
        if (hasMore) {
          log(index, 'table', 'insertParagraphAfter');
          cursor = table.insertParagraphAfter('');
          await doc.commit();
        } else {
          cursor = tableAnchor;
        }
        index++;
        continue;
      }

      if (block.type === 'listItem') {
        const { items, endIndex } = collectListRun(blocks, index, block.listIndex);
        await processListRun(items, index);
        index = endIndex;
        continue;
      }

      if (block.type === 'code') {
        // One paragraph per line, not one paragraph with embedded line
        // breaks, so each line is independently selectable/editable like
        // any other paragraph.
        for (const line of block.text.split('\n')) {
          log(index, 'code', 'insertParagraphAfter(line)');
          const linePara = nextParagraph(line);
          if (line) linePara.setWholeTextFormat({ bold: false, italic: false, fontName: MONOSPACE_FONT });
          trackNonList(linePara, () => applyBodyStyle(linePara, options));
          cursor = linePara;
        }
        await doc.commit();
        index++;
        continue;
      }

      // heading / paragraph / quote
      const first = block.inlines[0];
      const firstText = first ? first.text : '';
      log(index, block.type, 'insertParagraphAfter');
      const p = nextParagraph(firstText);
      if (first) p.setWholeTextFormat(runFormat(first));
      appendRemainingInlines(p, block.inlines);

      switch (block.type) {
        case 'heading':
          trackNonList(p, () => p.setStyleBuiltIn(headingStyleName(block.level)));
          break;
        case 'quote':
          trackNonList(p, () => p.setStyleBuiltIn('Quote'));
          break;
        case 'paragraph':
        default:
          trackNonList(p, () => applyBodyStyle(p, options));
          break;
      }
      await doc.commit();
      cursor = p;
      index++;
    } catch (err) {
      return fail(err);
    }
  }

  try {
    log(-1, 'verify', 'verifyAndRepairListItems');
    await doc.verifyAndRepairListItems(tracked);
  } catch (err) {
    // The paste itself already succeeded at this point; don't report a
    // failure over the safety net alone.
    console.warn('AI Paste: verifyAndRepairListItems failed', err);
  }
}

/** Gathers a contiguous run of listItem blocks sharing `listIndex`, starting at `startIndex` (nested sub-items share their parent's listIndex, so a whole list — nesting included — is always contiguous). */
function collectListRun(blocks: Block[], startIndex: number, listIndex: number): { items: ListItemBlock[]; endIndex: number } {
  const items: ListItemBlock[] = [];
  let i = startIndex;
  while (i < blocks.length) {
    const b = blocks[i];
    if (b.type === 'listItem' && b.listIndex === listIndex) {
      items.push(b);
      i++;
    } else {
      break;
    }
  }
  return { items, endIndex: i };
}

function applyBodyStyle(p: WordParagraphHandle, options: InsertOptions): void {
  if (options.bodyStyle.kind === 'custom') p.setStyleByName(options.bodyStyle.name);
  else p.setStyleBuiltIn('Normal');
}

function headingStyleName(level: number): string {
  const clamped = Math.min(Math.max(Math.round(level), 1), MAX_HEADING_LEVEL);
  return `Heading${clamped}`;
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
