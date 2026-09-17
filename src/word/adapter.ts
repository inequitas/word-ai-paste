/**
 * A small abstraction over the slice of Office.js we actually use.
 *
 * `insert.ts` is written entirely against this interface (never against the
 * `Word`/`Office` globals directly), so it can be unit-tested with a plain
 * in-memory fake — see tests/fakeWordAdapter.ts — instead of a real Word
 * runtime. `officeAdapter.ts` is the thin real implementation used by the
 * task pane.
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
  /** Appends one formatted run of text to the end of this paragraph. `text` may contain "\n" for a manual line break. */
  insertRun(text: string, format: RunFormat): void;
  /**
   * Creates a new empty paragraph after this one and returns it. NOTE: in
   * real Word, calling this on a paragraph that is currently a list item
   * usually creates *another list item* (the same thing that happens when
   * you press Enter at the end of a list line) — setting a style afterwards
   * does not reliably remove that list membership. Callers that want a
   * plain, non-list paragraph after a list must go through
   * `WordListHandle.insertParagraphAfter()` instead (see insert.ts), and
   * everything is additionally checked and repaired after the main sync —
   * see `WordDocument.verifyAndRepairListItems`.
   */
  insertParagraphAfter(): WordParagraphHandle;
  /** Creates a table after this paragraph with plain-text cell values and returns it. */
  insertTableAfter(values: string[][]): WordTableHandle;
  /** Starts a brand new Word list with this paragraph as its first (level 0) item. */
  startNewList(): WordListHandle;
  /** Marks this paragraph as level `level` of an already-open list (paragraph must already be a member — see WordListHandle.insertItemAfter). */
  setListLevel(level: number): void;
  /** Removes this paragraph from whatever list it belongs to, if any (no-op otherwise). */
  detachFromList(): void;
}

export interface WordListHandle {
  /** Appends a new empty paragraph as a member of this list (default level 0) and returns it. */
  insertItemAfter(): WordParagraphHandle;
  /**
   * Creates a new empty paragraph immediately after the *whole list* (not
   * a member of it) and returns it — the correct way to get back to plain
   * body content right after a list.
   */
  insertParagraphAfter(): WordParagraphHandle;
  /** Configures level `level` as a bullet level, once. Idempotent per level. */
  ensureBulletLevel(level: number): void;
  /** Configures level `level` as a simple "1." Arabic numbering level, once. Idempotent per level. */
  ensureNumberLevel(level: number, startAt?: number): void;
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
  applyStyle(settings: TableStyleSettings): void;
  /**
   * Creates and returns a brand new paragraph immediately after the table
   * (pushing anything that already followed it further down). Only call
   * this when more of our own blocks still need to be inserted after the
   * table — if the table is the last thing we insert, leave it alone and
   * rely on Word's own structural paragraph after a trailing table.
   */
  insertParagraphAfter(): WordParagraphHandle;
}

export interface CursorInfo {
  /** The (now-empty, reusable) paragraph at the insertion point. A non-empty selection is cleared first. */
  paragraph: WordParagraphHandle;
  /** True when that paragraph has no visible text, so the first inserted block can reuse it. */
  isEmpty: boolean;
  /** The list the cursor paragraph currently belongs to, if any. */
  list: WordListHandle | null;
}

/** One paragraph we inserted/touched, and whether it was meant to end up as a list item. */
export interface ParagraphIntent {
  handle: WordParagraphHandle;
  intendedListItem: boolean;
  /** Non-list paragraphs only: re-applies the style this paragraph is supposed to have, called again during repair. */
  reapplyStyle?: () => void;
}

export interface WordDocument {
  /** One-time read of the cursor position. Only method that may need a context.sync() internally. */
  getCursor(): Promise<CursorInfo>;
  /** Flushes all queued operations. Must be called exactly once, after all insertion calls. */
  commit(): Promise<void>;
  /**
   * Safety net for the "paragraph after a list stays a list item" Word
   * quirk: loads `isListItem` for every tracked paragraph, syncs, then for
   * each paragraph that was NOT meant to be a list item but is one anyway,
   * detaches it and re-applies its intended style (syncing once more).
   * Paragraphs that were meant to be list items but aren't are only
   * logged via console.warn — not auto-fixed, since we can't safely guess
   * which list they should have joined.
   */
  verifyAndRepairListItems(records: ParagraphIntent[]): Promise<void>;
}
