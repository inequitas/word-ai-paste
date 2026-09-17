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
  /** Creates a new empty paragraph after this one and returns it. */
  insertParagraphAfter(): WordParagraphHandle;
  /** Creates a table after this paragraph with plain-text cell values and returns it. */
  insertTableAfter(values: string[][]): WordTableHandle;
  /** Starts a brand new Word list with this paragraph as its first (level 0) item. */
  startNewList(): WordListHandle;
  /** Marks this paragraph as level `level` of an already-open list (paragraph must already be a member — see WordListHandle.insertItemAfter). */
  setListLevel(level: number): void;
}

export interface WordListHandle {
  /** Appends a new empty paragraph as a member of this list (default level 0) and returns it. */
  insertItemAfter(): WordParagraphHandle;
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
  /** The paragraph at the insertion point (or the start of a non-empty selection). */
  paragraph: WordParagraphHandle;
  /** True when that paragraph has no visible text, so the first inserted block can reuse it. */
  isEmpty: boolean;
}

export interface WordDocument {
  /** One-time read of the cursor position. Only method that may need a context.sync() internally. */
  getCursor(): Promise<CursorInfo>;
  /** Flushes all queued operations. Must be called exactly once, after all insertion calls. */
  commit(): Promise<void>;
}
