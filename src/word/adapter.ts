/**
 * A small abstraction over the slice of Office.js we actually use.
 *
 * `insert.ts` is written entirely against this interface (never against the
 * `Word`/`Office` globals directly), so it can be unit-tested with a plain
 * in-memory fake — see tests/fakeWordAdapter.ts — instead of a real Word
 * runtime. `officeAdapter.ts` is the thin real implementation used by the
 * task pane.
 *
 * Execution model: step-wise, not one big batch. Methods returning `void`
 * only QUEUE work; `insert.ts` calls `WordDocument.commit()` after each
 * block / list step so a failure points at one step. Methods returning a
 * Promise do their own load + sync.
 *
 * Deliberately absent: `Word.List.insertParagraph`. On Word for Mac
 * (16.113) `list.insertParagraph(text, "End")` does not add the paragraph
 * to the list — it lands directly after the list's current range (which
 * stays just the first item), so repeated calls come out in reverse order —
 * and `"After"` behaves the same way. See insert.ts for the two-phase
 * design that replaced it.
 */

export interface RunFormat {
  bold: boolean;
  italic: boolean;
  /** Font family to force (used only for code runs); omit to inherit from the style. */
  fontName?: string;
  link?: string;
}

export interface WordParagraphHandle {
  /** Sets a locale-independent built-in style (e.g. "Heading1", "Normal", "ListParagraph", "Quote"). */
  setStyleBuiltIn(name: string): void;
  /** Sets a custom or localized style by its display name (document-specific, not portable). */
  setStyleByName(name: string): void;
  /** Appends one more formatted run after whatever text this paragraph already has. */
  insertRun(text: string, format: RunFormat): void;
  /**
   * Applies bold/italic/font/link formatting to this paragraph's entire
   * current text. Meant to be called right after creating the paragraph
   * WITH its first run's text already passed to `insertParagraphAfter`.
   * No-op-safe on a genuinely empty paragraph (a blank spacer).
   */
  setWholeTextFormat(format: RunFormat): void;
  /**
   * Creates a new paragraph directly after this one, with `text` as its
   * content from the moment of creation (pass "" for an empty spacer
   * paragraph), and returns it. This is the only way insert.ts creates
   * paragraphs next to other paragraphs; it keeps document order and
   * styles on Word for Mac.
   *
   * NOTE: if THIS paragraph is a list item, real Word may make the new
   * paragraph a list item too (like pressing Enter at the end of a list
   * line). insert.ts only ever calls this on paragraphs that are not list
   * items yet (all lists are created afterwards, in phase 2), except for a
   * cursor paragraph that already was one — handled explicitly there.
   */
  insertParagraphAfter(text: string): WordParagraphHandle;
  /** Creates a table after this paragraph with plain-text cell values and returns it. */
  insertTableAfter(values: string[][]): WordTableHandle;
  /** Starts a brand new Word list with this paragraph as its first (level 0) item. */
  startNewList(): WordListHandle;
  /**
   * Makes this (plain, non-list) paragraph an item of the existing list
   * `listId` at `level`. Word fails this call if the paragraph already is
   * a list item.
   */
  attachToList(listId: number, level: number): void;
  /**
   * Sets this list item's level. Only call it AFTER a commit that followed
   * the paragraph joining a list (it uses the throwing `listItem` accessor).
   */
  setListLevel(level: number): void;
  /**
   * Removes this paragraph from its list, if any. Only call this on a
   * paragraph whose list-item status was confirmed by a prior read.
   */
  detachFromList(): void;
}

export interface WordListHandle {
  /** Loads and returns this list's numeric id (own sync). */
  getId(): Promise<number>;
  /** Queues: level `level` shows a solid bullet. */
  setLevelBullet(level: number): void;
  /** Queues: level `level` shows simple Arabic numbering, "1.". */
  setLevelNumbering(level: number): void;
  /**
   * Queues `List.setLevelIndents(level, textIndentPt, bulletIndentPt)`.
   * insert.ts passes the geometry of Word's own bullet button — see
   * `listLevelIndents` there.
   */
  setLevelIndents(level: number, textIndentPt: number, bulletIndentPt: number): void;
  /** Queues: level `level` starts counting at `startingNumber`. */
  setLevelStartingNumber(level: number, startingNumber: number): void;
}

export interface TableStyleSettings {
  /** A portable built-in table style name (e.g. "GridTable4"), or null to use `customName`. */
  styleBuiltIn: string | null;
  /** A custom/document-defined table style display name, used only when `styleBuiltIn` is null. */
  customName: string | null;
  headerRowCount: 0 | 1;
  styleFirstColumn: boolean;
  styleBandedRows: boolean;
  styleBandedColumns: boolean;
  styleLastColumn: boolean;
  styleTotalRow: boolean;
}

export interface WordTableHandle {
  /**
   * Syncs (the table must exist server-side before it can be styled),
   * then applies `settings`. Never throws: if `styleBuiltIn` fails, falls
   * back to setting `style` by its English display name; if that also
   * fails, logs a warning and leaves the table with Word's default table
   * look rather than failing the whole insert.
   */
  applyStyle(settings: TableStyleSettings): Promise<void>;
  /**
   * Creates and returns a new paragraph (with `text` from creation)
   * directly after the table. insert.ts only calls this when another of
   * our own blocks follows the table.
   */
  insertParagraphAfter(text: string): WordParagraphHandle;
}

export interface CursorInfo {
  /** The paragraph at the insertion point. A non-empty selection is deleted first. */
  paragraph: WordParagraphHandle;
  /** True when that paragraph has no visible text, so the first inserted block can reuse it. */
  isEmpty: boolean;
  /** Whether that paragraph is currently a list item (confirmed via a load, never guessed). */
  isListItem: boolean;
}

/** One paragraph we inserted/touched, and whether it was meant to end up as a list item. */
export interface ParagraphIntent {
  handle: WordParagraphHandle;
  intendedListItem: boolean;
  /** Non-list paragraphs only: re-applies the style this paragraph is supposed to have, called again during repair. */
  reapplyStyle?: () => void;
}

/** List state of one paragraph, as read back from the document. */
export interface ListState {
  isListItem: boolean;
  /** Id of the list the paragraph belongs to, or null (not a list item, or not readable). */
  listId: number | null;
  /** The paragraph's list level, or null. */
  level: number | null;
}

/** One paragraph as read back from the document, in document order. */
export interface ReadBackParagraph {
  text: string;
  isListItem: boolean;
  /** `Paragraph.styleBuiltIn` ("Other" for a custom style). */
  styleBuiltIn: string;
  leftIndent: number | null;
  firstLineIndent: number | null;
  listId: number | null;
  level: number | null;
  /** The rendered bullet/number, e.g. "1." — null if not a list item or unreadable. */
  listString: string | null;
}

export interface ReadBackTable {
  rowCount: number;
  headerRowCount: number;
}

export interface ReadBackResult {
  /** Body paragraphs from the first to the last given paragraph, in document order (table-cell paragraphs excluded). */
  paragraphs: ReadBackParagraph[];
  /** Tables inside that range. */
  tables: ReadBackTable[];
  /** Anything that could not be read (e.g. list ids), for the step log. */
  notes: string[];
}

export interface WordDocument {
  /** Reads the cursor position. Deletes a non-empty selection first (never inserts "" — see officeAdapter). */
  getCursor(): Promise<CursorInfo>;
  /** Flushes all queued operations. */
  commit(): Promise<void>;
  /** Reads list membership, list id and level of each paragraph (one sync). */
  readListState(paragraphs: WordParagraphHandle[]): Promise<ListState[]>;
  /**
   * Safety net: loads `isListItem` for every tracked paragraph and, for
   * each one that was NOT meant to be a list item but is one anyway,
   * detaches it and re-applies its intended style. Returns how many
   * paragraphs it repaired. Missing list items are reported by the
   * read-back (phase 3), not here.
   */
  verifyAndRepairListItems(records: ParagraphIntent[]): Promise<number>;
  /**
   * Reads back every body paragraph from `first` to `last` (inclusive), in
   * document order, plus the tables in between.
   */
  readBack(first: WordParagraphHandle, last: WordParagraphHandle): Promise<ReadBackResult>;
}
