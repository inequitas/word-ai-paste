/**
 * A small abstraction over the slice of Office.js we actually use.
 *
 * `insert.ts` is written entirely against this interface (never against the
 * `Word`/`Office` globals directly), so it can be unit-tested with a plain
 * in-memory fake — see tests/fakeWordAdapter.ts — instead of a real Word
 * runtime. `officeAdapter.ts` is the thin real implementation used by the
 * task pane.
 *
 * Execution model: step-wise, not one big batch. `insert.ts` calls
 * `WordDocument.commit()` after each block (occasionally twice — see
 * `WordParagraphHandle.setListLevel`), so a failure points at one block and
 * everything before it has already taken effect. A few operations that are
 * "nice to have but not worth failing the whole paste over" — configuring a
 * list level's bullet/number format, applying a table style — do their own
 * internal sync, catch their own errors, and simply log and continue rather
 * than throwing; see their doc comments below.
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
   * WITH its first run's text already passed to `insertParagraphAfter` /
   * `WordListHandle.insertItemAfter` — i.e. instead of ever creating an
   * empty paragraph and then filling it, the first run's text is there
   * from the moment of creation, and this just formats it. No-op-safe on
   * a genuinely empty paragraph (a blank spacer).
   */
  setWholeTextFormat(format: RunFormat): void;
  /**
   * Creates a new paragraph after this one, with `text` as its content
   * from the moment of creation (pass "" for a genuinely empty spacer
   * paragraph), and returns it.
   *
   * NOTE: in real Word, calling this on a paragraph that is currently a
   * list item usually creates *another list item* (the same thing that
   * happens when you press Enter at the end of a list line) — setting a
   * style afterwards does not reliably remove that list membership.
   * Callers that want a plain, non-list paragraph after a list must go
   * through `WordListHandle.insertParagraphAfter()` instead (see
   * insert.ts), and everything is additionally checked and repaired after
   * the main sync — see `WordDocument.verifyAndRepairListItems`.
   */
  insertParagraphAfter(text: string): WordParagraphHandle;
  /** Creates a table after this paragraph with plain-text cell values and returns it. */
  insertTableAfter(values: string[][]): WordTableHandle;
  /**
   * Starts a brand new Word list with this paragraph as its first (level
   * 0) item. Only call this on a paragraph that already has its text (see
   * `insertParagraphAfter`) — starting a list on a still-empty paragraph
   * is suspected of being one cause of Word throwing a bare
   * `GeneralException` on this operation.
   */
  startNewList(): WordListHandle;
  /**
   * Marks this paragraph as level `level` of an already-open list
   * (paragraph must already be a confirmed member — see
   * `WordListHandle.insertItemAfter`). Only call this AFTER a
   * `WordDocument.commit()` that happened after the paragraph joined the
   * list — setting the level in the same unsynced batch as the join is
   * the other suspected `GeneralException` cause.
   */
  setListLevel(level: number): void;
  /**
   * Removes this paragraph from whatever list it belongs to, if any
   * (no-op otherwise). Only call this on a paragraph whose list-item
   * status was actually confirmed by a prior load+sync (never
   * speculatively) — see insert.ts's callers.
   */
  detachFromList(): void;
}

export interface WordListHandle {
  /** Appends a new paragraph as a member of this list (default level 0), with `text` as its content from creation, and returns it. */
  insertItemAfter(text: string): WordParagraphHandle;
  /**
   * Creates a new paragraph immediately after the *whole list* (not a
   * member of it), with `text` as its content from creation, and returns
   * it — the correct way to get back to plain body content right after a
   * list.
   */
  insertParagraphAfter(text: string): WordParagraphHandle;
  /**
   * Configures level `level` as a bullet level, once (idempotent per
   * level). Syncs and catches its own errors: if Word rejects the
   * formatting call, this logs a warning and resolves anyway, leaving the
   * level at Word's default bullet/number formatting rather than failing
   * the whole insert over a cosmetic setting.
   */
  ensureBulletLevel(level: number): Promise<void>;
  /** Same as `ensureBulletLevel`, but for a simple "1." Arabic numbering level. */
  ensureNumberLevel(level: number, startAt?: number): Promise<void>;
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
   * Creates and returns a brand new paragraph immediately after the table
   * (pushing anything that already followed it further down). Only call
   * this when more of our own blocks still need to be inserted after the
   * table — if the table is the last thing we insert, leave it alone and
   * rely on Word's own structural paragraph after a trailing table.
   */
  insertParagraphAfter(text: string): WordParagraphHandle;
}

export interface CursorInfo {
  /** The (now-empty, reusable) paragraph at the insertion point. A non-empty selection is deleted first. */
  paragraph: WordParagraphHandle;
  /** True when that paragraph has no visible text, so the first inserted block can reuse it. */
  isEmpty: boolean;
  /** The list the cursor paragraph currently belongs to, if any (confirmed via a load, never guessed). */
  list: WordListHandle | null;
}

/** One paragraph we inserted/touched, and whether it was meant to end up as a list item. */
export interface ParagraphIntent {
  handle: WordParagraphHandle;
  intendedListItem: boolean;
  /** Non-list paragraphs only: re-applies the style this paragraph is supposed to have, called again during repair. */
  reapplyStyle?: () => void;
}

/** One list item for the `insertOoxmlList` fallback (see src/word/ooxml.ts). */
export interface OoxmlListItemInput {
  level: number;
  ordered: boolean;
  inlines: import('../model').Inline[];
}

export interface WordDocument {
  /** Reads the cursor position. Deletes a non-empty selection first (never inserts "" — see officeAdapter). */
  getCursor(): Promise<CursorInfo>;
  /** Flushes all queued operations. Called after every block (and once more after a nested list item's level is set). */
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
  /**
   * Fallback for when the paragraph/list API path throws while starting or
   * continuing a list: inserts the whole `items` run at once via raw OOXML
   * (see src/word/ooxml.ts). Replaces `anchor` if `anchorIsEmpty`,
   * otherwise inserts after it. Syncs internally. Returns a handle per
   * item (already styled "ListParagraph") and — best-effort — the real
   * `Word.List` those items ended up on, so anything that follows can
   * still use the normal list-escape logic.
   */
  insertOoxmlList(
    anchor: WordParagraphHandle,
    anchorIsEmpty: boolean,
    items: OoxmlListItemInput[]
  ): Promise<{ items: WordParagraphHandle[]; list: WordListHandle | null }>;
}
