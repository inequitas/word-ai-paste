import type { Block, Inline, TableBlock } from '../model';
import type { WordDocument, WordListHandle, WordParagraphHandle, TableStyleSettings, ParagraphIntent } from './adapter';

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

/**
 * List-numbering-restart strategy
 * --------------------------------
 * Every distinct `listIndex` produced by the parsers gets its own
 * `Word.List` object, created with `paragraph.startNewList()` for that
 * list's first item. Every later item of the same list — nested sub-items
 * included — is appended with `list.insertParagraph(..., "End")` and then,
 * for anything below the top level, `paragraph.listItemOrNullObject.level`
 * is set explicitly.
 *
 * This is deliberately different from the more obvious-looking
 * `paragraph.attachToList(listId, level)` route: `attachToList` needs the
 * *numeric* list id as a plain argument, and reading `list.id` back off a
 * freshly created `Word.List` proxy requires its own `load()` +
 * `context.sync()` before you can use it — exactly the extra round trip
 * this module is trying to avoid (the goal is one load + one final sync
 * for the whole paste). Holding on to the in-memory `Word.List` object
 * returned by `startNewList()` and calling its own `insertParagraph`
 * sidesteps that: no id ever needs to be read back, and a brand new
 * `Word.List` restarts numbering at 1 by construction, which is exactly
 * the "each separate numbered list starts at 1" requirement.
 *
 * All of `startNewList`, `List.insertParagraph`, `setLevelBullet`,
 * `setLevelNumbering`, `setLevelStartingNumber` and
 * `ListItem.level` are WordApi 1.3 — the same minimum the manifest already
 * requires for `styleBuiltIn` itself, so there is no lower-tier fallback
 * that would still let us apply built-in styles by their portable,
 * locale-independent name. A host that lacks WordApi 1.3 therefore can't
 * safely run this add-in at all; `taskpane.ts` checks
 * `Office.context.requirements.isSetSupported('WordApi', '1.3')` once at
 * startup and shows a plain "please update Word" message instead of
 * attempting a degraded insert. The one thing only a real-Word test can
 * settle is whether a level's bullet/number counter set via
 * `setLevelBullet`/`setLevelNumbering` truly behaves independently per
 * nesting level the way Word's own bullet button does, for deeply nested
 * lists.
 *
 * Leaving a list cleanly
 * -----------------------
 * In real Word, `paragraph.insertParagraphAfter()` on a paragraph that is
 * currently a list item usually creates *another list item* — the same
 * thing that happens when you press Enter at the end of a list line — and
 * setting `styleBuiltIn` afterwards does not reliably remove that list
 * membership (the style and the list/numPr are separate paragraph
 * properties). Since almost every list in Kevin's documents is followed by
 * body text, a blank paragraph, or a heading, this module tracks whichever
 * `Word.List` the "current position" belongs to (`currentList`) and, for
 * every block that is *not* a list item, gets its paragraph via
 * `list.insertParagraphAfter()` instead — which inserts after the whole
 * list rather than as one more member of it — then forgets `currentList`.
 * Tables are the one exception: they aren't paragraphs, so inserting one
 * directly after a list-item paragraph doesn't carry any numPr along with
 * it.
 *
 * As a safety net on top of that (for anything this reasoning gets wrong,
 * and for the "reused an already-empty bullet as the cursor" case), every
 * non-list paragraph we create is tracked and re-checked with
 * `WordDocument.verifyAndRepairListItems` after the main sync — see that
 * method's doc comment.
 */

export async function insertBlocks(
  doc: WordDocument,
  blocks: Block[],
  options: InsertOptions = DEFAULT_INSERT_OPTIONS
): Promise<void> {
  if (!blocks.length) return;

  const { paragraph: anchor, isEmpty, list: initialList } = await doc.getCursor();
  let cursor = anchor;
  let anchorReusable = isEmpty;
  let currentList: WordListHandle | null = initialList;
  const listHandles = new Map<number, WordListHandle>();
  const tracked: ParagraphIntent[] = [];

  /** A fresh paragraph that is guaranteed not to be a member of any list. */
  const paragraphOutsideList = (): WordParagraphHandle => {
    const p = currentList ? currentList.insertParagraphAfter() : cursor.insertParagraphAfter();
    currentList = null;
    return p;
  };

  const nextParagraph = (): WordParagraphHandle => {
    if (anchorReusable) {
      anchorReusable = false;
      if (currentList) {
        // The reused cursor paragraph was itself an (empty) list item.
        cursor.detachFromList();
        currentList = null;
      }
      return cursor;
    }
    return paragraphOutsideList();
  };

  const trackNonList = (p: WordParagraphHandle, reapplyStyle: () => void): void => {
    reapplyStyle();
    tracked.push({ handle: p, intendedListItem: false, reapplyStyle });
  };

  blocks.forEach((block, index) => {
    if (block.type === 'table') {
      // Not a paragraph, so it never inherits list numbering — safe to
      // insert straight after the current cursor paragraph either way.
      const tableAnchor = cursor;
      anchorReusable = false;
      currentList = null;
      const table = tableAnchor.insertTableAfter(tableValues(block));
      const headerRowCount = block.header ? options.tableStyle.headerRowCount : 0;
      table.applyStyle({ ...options.tableStyle, headerRowCount });
      const hasMore = index < blocks.length - 1;
      cursor = hasMore ? table.insertParagraphAfter() : tableAnchor;
      return;
    }

    if (block.type === 'listItem') {
      let list = listHandles.get(block.listIndex);
      let p: WordParagraphHandle;
      if (!list) {
        p = nextParagraph(); // guaranteed not already a list item, as startNewList() requires
        list = p.startNewList();
        listHandles.set(block.listIndex, list);
      } else {
        p = list.insertItemAfter();
      }
      currentList = list;
      p.setStyleBuiltIn('ListParagraph');
      if (block.level > 0) p.setListLevel(block.level);
      if (block.ordered) list.ensureNumberLevel(block.level, block.start);
      else list.ensureBulletLevel(block.level);
      insertInlines(p, block.inlines);
      tracked.push({ handle: p, intendedListItem: true });
      cursor = p;
      return;
    }

    if (block.type === 'code') {
      // One paragraph per line, not one paragraph with embedded line
      // breaks, so each line is independently selectable/editable like any
      // other paragraph.
      for (const line of block.text.split('\n')) {
        const linePara = nextParagraph();
        trackNonList(linePara, () => applyBodyStyle(linePara, options));
        linePara.insertRun(line, { bold: false, italic: false, fontName: MONOSPACE_FONT });
        cursor = linePara;
      }
      return;
    }

    const p = nextParagraph();
    switch (block.type) {
      case 'heading':
        trackNonList(p, () => p.setStyleBuiltIn(headingStyleName(block.level)));
        insertInlines(p, block.inlines);
        break;
      case 'quote':
        trackNonList(p, () => p.setStyleBuiltIn('Quote'));
        insertInlines(p, block.inlines);
        break;
      case 'paragraph':
      default:
        trackNonList(p, () => applyBodyStyle(p, options));
        insertInlines(p, block.inlines);
        break;
    }
    cursor = p;
  });

  await doc.commit();
  await doc.verifyAndRepairListItems(tracked);
}

function applyBodyStyle(p: WordParagraphHandle, options: InsertOptions): void {
  if (options.bodyStyle.kind === 'custom') p.setStyleByName(options.bodyStyle.name);
  else p.setStyleBuiltIn('Normal');
}

function headingStyleName(level: number): string {
  const clamped = Math.min(Math.max(Math.round(level), 1), MAX_HEADING_LEVEL);
  return `Heading${clamped}`;
}

function insertInlines(p: WordParagraphHandle, inlines: Inline[]): void {
  for (const inline of inlines) {
    if (!inline.text) continue;
    p.insertRun(inline.text, {
      bold: !!inline.bold,
      italic: !!inline.italic,
      fontName: inline.code ? MONOSPACE_FONT : undefined,
      link: inline.link
    });
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
