import type {
  CursorInfo,
  ListState,
  ParagraphIntent,
  ReadBackParagraph,
  ReadBackResult,
  RunFormat,
  TableStyleSettings,
  WordDocument,
  WordListHandle,
  WordParagraphHandle,
  WordTableHandle
} from '../src/word/adapter';

/**
 * A small in-memory implementation of the Word adapter interfaces, used to
 * unit-test src/word/insert.ts without a real Word runtime.
 *
 * Real-Word behaviour it reproduces on purpose:
 * - `insertParagraphAfter()` on a paragraph that is a list item creates
 *   another item of the same list/level (like pressing Enter at the end of
 *   a list line).
 * - The Word-for-Mac 16.113 quirk: `FakeList.insertParagraph(text, "End" |
 *   "After")` inserts directly after the list's FIRST paragraph and does not
 *   join the list, so repeated calls come out in reverse order. It is not
 *   part of the adapter interface (insert.ts must never use it); tests
 *   call it directly to show the fake matches what Kevin saw in Word.
 * - `attachToList()` throws for a paragraph that already is a list item.
 * - Numbers in `listString` follow the list's members in document order.
 *
 * Failure injection: `failNext(action, times, skip)` makes calls to that
 * action throw an Office.js-shaped error (after `skip` successful ones);
 * `attachToListIsNoop = true`
 * makes `attachToList` silently do nothing (isListItem stays false).
 */

export interface FakeRun {
  text: string;
  bold: boolean;
  italic: boolean;
  fontName?: string;
  link?: string;
}

export interface FakeParagraph {
  kind: 'paragraph';
  styleBuiltIn?: string;
  styleName?: string;
  /** One array per line — a hard line break starts a new line array. */
  lines: FakeRun[][];
  listRef?: { list: FakeList; level: number };
  /** Explicit left indent in points, set by setIndents(). */
  leftIndent?: number;
  /** Explicit first-line indent in points, set by setIndents(). */
  firstLineIndent?: number;
}

export interface FakeTable {
  kind: 'table';
  values: string[][];
  style: TableStyleSettings | null;
}

export type FakeBlock = FakeParagraph | FakeTable;

export function paragraphText(p: FakeParagraph): string {
  return p.lines.map((line) => line.map((r) => r.text).join('')).join('\n');
}

export function isParagraph(b: FakeBlock): b is FakeParagraph {
  return b.kind === 'paragraph';
}

export function isTable(b: FakeBlock): b is FakeTable {
  return b.kind === 'table';
}

/** Shape of the fake error thrown to simulate Word's GeneralException, for tests to inspect. */
export class FakeOfficeError extends Error {
  code = 'GeneralException';
  debugInfo: { errorLocation: string; statement: string };
  constructor(where: string) {
    super('GeneralException');
    this.debugInfo = { errorLocation: where, statement: where };
  }
}

class FakeDocumentModel {
  blocks: FakeBlock[] = [];
  lists: FakeList[] = [];
}

function insertAfter(model: FakeDocumentModel, anchor: FakeBlock, block: FakeBlock): void {
  const idx = model.blocks.indexOf(anchor);
  if (idx === -1) throw new Error('fake adapter: anchor block not found');
  model.blocks.splice(idx + 1, 0, block);
}

function emptyParagraph(): FakeParagraph {
  return { kind: 'paragraph', lines: [[]] };
}

function paragraphWithText(text: string): FakeParagraph {
  const p = emptyParagraph();
  if (text) appendText(p, text, { bold: false, italic: false });
  return p;
}

function appendText(p: FakeParagraph, text: string, format: RunFormat): void {
  const segments = text.split('\n');
  segments.forEach((seg, i) => {
    if (i > 0) p.lines.push([]);
    if (seg) {
      p.lines[p.lines.length - 1].push({
        text: seg,
        bold: format.bold,
        italic: format.italic,
        fontName: format.fontName,
        link: format.link
      });
    }
  });
}

export class FakeParagraphHandle implements WordParagraphHandle {
  constructor(
    private readonly model: FakeDocumentModel,
    private readonly doc: FakeWordDocument,
    readonly para: FakeParagraph
  ) {}

  setStyleBuiltIn(name: string): void {
    this.para.styleBuiltIn = name;
    this.para.styleName = undefined;
  }

  setStyleByName(name: string): void {
    this.para.styleName = name;
    this.para.styleBuiltIn = undefined;
  }

  insertRun(text: string, format: RunFormat): void {
    appendText(this.para, text, format);
  }

  setWholeTextFormat(format: RunFormat): void {
    for (const line of this.para.lines) {
      for (const run of line) {
        run.bold = format.bold;
        run.italic = format.italic;
        if (format.fontName) run.fontName = format.fontName;
        if (format.link) run.link = format.link;
      }
    }
  }

  insertParagraphAfter(text: string): WordParagraphHandle {
    this.doc.fire('insertParagraphAfter');
    const p = paragraphWithText(text);
    insertAfter(this.model, this.para, p);
    // Mirrors real Word: a paragraph inserted right after a list item
    // becomes another item of that same list/level.
    if (this.para.listRef) this.para.listRef.list.join(p, this.para.listRef.level);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }

  insertTableAfter(values: string[][]): WordTableHandle {
    const t: FakeTable = { kind: 'table', values, style: null };
    insertAfter(this.model, this.para, t);
    return new FakeTableHandle(this.model, this.doc, t);
  }

  startNewList(): FakeList {
    this.doc.fire('startNewList');
    if (this.para.listRef) this.para.listRef.list.leave(this.para);
    const list = new FakeList(this.model, this.doc);
    list.join(this.para, 0);
    return list;
  }

  attachToList(listId: number, level: number): void {
    this.doc.fire('attachToList');
    if (this.doc.attachToListIsNoop) return;
    if (this.para.listRef) throw new FakeOfficeError('Word.Paragraph.attachToList (already a list item)');
    const list = this.model.lists.find((l) => l.id === listId);
    if (!list) throw new FakeOfficeError(`Word.Paragraph.attachToList (no list ${listId})`);
    list.join(this.para, level);
  }

  setListLevel(level: number): void {
    this.doc.fire('setListLevel');
    if (!this.para.listRef) throw new FakeOfficeError('Word.Paragraph.listItem (not a list item)');
    this.para.listRef.level = level;
  }

  detachFromList(): void {
    if (this.para.listRef) this.para.listRef.list.leave(this.para);
  }

  setIndents(leftIndentPt: number, firstLineIndentPt: number): void {
    this.para.leftIndent = leftIndentPt;
    this.para.firstLineIndent = firstLineIndentPt;
  }
}

let nextListId = 1000;

export class FakeList implements WordListHandle {
  readonly id = nextListId++;
  /** Member paragraphs, kept in document order. */
  members: FakeParagraph[] = [];
  readonly bulletLevels = new Set<number>();
  readonly numberLevels = new Set<number>();
  readonly startingNumbers = new Map<number, number>();
  readonly levelIndents = new Map<number, { textIndent: number; bulletIndent: number }>();

  constructor(
    private readonly model: FakeDocumentModel,
    private readonly doc: FakeWordDocument
  ) {
    model.lists.push(this);
  }

  join(p: FakeParagraph, level: number): void {
    p.listRef = { list: this, level };
    this.members.push(p);
    this.members.sort((a, b) => this.model.blocks.indexOf(a) - this.model.blocks.indexOf(b));
  }

  leave(p: FakeParagraph): void {
    const idx = this.members.indexOf(p);
    if (idx !== -1) this.members.splice(idx, 1);
    p.listRef = undefined;
  }

  /**
   * NOT part of WordListHandle. Reproduces Word for Mac: the new paragraph
   * lands directly after the list's first paragraph and is not a member.
   */
  insertParagraph(text: string, _location: 'End' | 'After'): FakeParagraphHandle {
    const p = paragraphWithText(text);
    insertAfter(this.model, this.members[0], p);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }

  async getId(): Promise<number> {
    this.doc.fire('getId');
    return this.id;
  }

  setLevelBullet(level: number): void {
    this.doc.fire('setLevelBullet');
    this.bulletLevels.add(level);
    this.numberLevels.delete(level);
  }

  setLevelNumbering(level: number): void {
    this.doc.fire('setLevelNumbering');
    this.numberLevels.add(level);
    this.bulletLevels.delete(level);
  }

  setLevelIndents(level: number, textIndent: number, bulletIndentRelative: number): void {
    this.doc.fire('setLevelIndents');
    // Store the absolute position: relative indent is applied to textIndent
    const bulletIndentAbsolute = textIndent + bulletIndentRelative;
    this.levelIndents.set(level, { textIndent, bulletIndent: bulletIndentAbsolute });
  }

  setLevelStartingNumber(level: number, startingNumber: number): void {
    this.doc.fire('setLevelStartingNumber');
    this.startingNumbers.set(level, startingNumber);
  }

  /** The bullet/number Word would show for `p`. */
  listString(p: FakeParagraph): string {
    const level = p.listRef!.level;
    if (!this.numberLevels.has(level)) return '•';
    let n = (this.startingNumbers.get(level) ?? 1) - 1;
    for (const m of this.members) {
      const l = m.listRef!.level;
      if (l < level) n = (this.startingNumbers.get(level) ?? 1) - 1;
      if (l === level) n++;
      if (m === p) break;
    }
    return `${n}.`;
  }
}

class FakeTableHandle implements WordTableHandle {
  constructor(
    private readonly model: FakeDocumentModel,
    private readonly doc: FakeWordDocument,
    readonly table: FakeTable
  ) {}

  async applyStyle(settings: TableStyleSettings): Promise<void> {
    this.table.style = settings;
  }

  insertParagraphAfter(text: string): WordParagraphHandle {
    const p = paragraphWithText(text);
    insertAfter(this.model, this.table, p);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }
}

export type FailableAction =
  | 'insertParagraphAfter'
  | 'startNewList'
  | 'attachToList'
  | 'setListLevel'
  | 'getId'
  | 'setLevelBullet'
  | 'setLevelNumbering'
  | 'setLevelIndents'
  | 'setLevelStartingNumber'
  | 'readBack';

/** List Paragraph style's own indent, used for a ListParagraph paragraph that is not (yet) a list item. */
const LIST_PARAGRAPH_STYLE_INDENT = 36;

export class FakeWordDocument implements WordDocument {
  readonly model = new FakeDocumentModel();
  readonly anchor: FakeParagraph;
  commitCount = 0;
  /** How many paragraphs the repair pass detached + restyled. */
  repairedCount = 0;
  /** When true, attachToList() silently does nothing. */
  attachToListIsNoop = false;
  /** Every call to a failable action, in order. */
  readonly calls: FailableAction[] = [];

  private failing = new Map<FailableAction, { skip: number; times: number }>();

  /**
   * @param options.anchorText   text of the cursor paragraph ("" = empty, reusable)
   * @param options.anchorInList the cursor paragraph is an item of an existing list
   * @param options.trailingText existing document content after the cursor paragraph
   */
  constructor(options: { anchorText?: string; anchorInList?: boolean; trailingText?: string } = {}) {
    this.anchor = paragraphWithText(options.anchorText ?? '');
    this.model.blocks.push(this.anchor);
    if (options.anchorInList) {
      new FakeList(this.model, this).join(this.anchor, 0);
    }
    if (options.trailingText !== undefined) {
      this.model.blocks.push(paragraphWithText(options.trailingText));
    }
  }

  /** Test hook: after `skip` successful calls, make the next `times` calls to this action throw a FakeOfficeError. */
  failNext(action: FailableAction, times = 1, skip = 0): void {
    this.failing.set(action, { skip, times });
  }

  /** Internal: records the call and throws if a failure was requested for it. */
  fire(action: FailableAction): void {
    this.calls.push(action);
    const plan = this.failing.get(action);
    if (!plan) return;
    if (plan.skip > 0) {
      plan.skip--;
      return;
    }
    if (plan.times > 0) {
      plan.times--;
      throw new FakeOfficeError(`Word.${action}`);
    }
  }

  /** Every paragraph's text, in document order. */
  texts(): string[] {
    return this.paragraphs.map(paragraphText);
  }

  get paragraphs(): FakeParagraph[] {
    return this.model.blocks.filter(isParagraph);
  }

  async getCursor(): Promise<CursorInfo> {
    const isEmpty = this.anchor.lines.length === 1 && this.anchor.lines[0].length === 0;
    return {
      paragraph: new FakeParagraphHandle(this.model, this, this.anchor),
      isEmpty,
      isListItem: Boolean(this.anchor.listRef)
    };
  }

  async commit(): Promise<void> {
    this.commitCount++;
  }

  async readListState(paragraphs: WordParagraphHandle[]): Promise<ListState[]> {
    return paragraphs.map((h) => {
      const ref = (h as FakeParagraphHandle).para.listRef;
      return { isListItem: Boolean(ref), listId: ref ? ref.list.id : null, level: ref ? ref.level : null };
    });
  }

  async verifyAndRepairListItems(records: ParagraphIntent[]): Promise<number> {
    let repaired = 0;
    for (const r of records) {
      const para = (r.handle as FakeParagraphHandle).para;
      if (!r.intendedListItem && para.listRef) {
        r.handle.detachFromList();
        r.reapplyStyle?.();
        repaired++;
      }
    }
    this.repairedCount += repaired;
    return repaired;
  }

  async readBack(first: WordParagraphHandle, last: WordParagraphHandle): Promise<ReadBackResult> {
    this.fire('readBack');
    const from = this.model.blocks.indexOf((first as FakeParagraphHandle).para);
    const to = this.model.blocks.indexOf((last as FakeParagraphHandle).para);
    if (from === -1 || to === -1 || to < from) throw new Error('fake adapter: bad read-back range');
    const slice = this.model.blocks.slice(from, to + 1);
    return {
      paragraphs: slice.filter(isParagraph).map((p) => readBackOf(p)),
      tables: slice.filter(isTable).map((t) => ({ rowCount: t.values.length, headerRowCount: t.style?.headerRowCount ?? 0 })),
      notes: []
    };
  }
}

export function readBackOf(p: FakeParagraph): ReadBackParagraph {
  const ref = p.listRef;
  // Check for explicit indents set by setIndents() first
  let leftIndent = p.leftIndent !== undefined ? p.leftIndent : (p.styleBuiltIn === 'ListParagraph' ? LIST_PARAGRAPH_STYLE_INDENT : 0);
  let firstLineIndent = p.firstLineIndent !== undefined ? p.firstLineIndent : 0;
  if (ref && p.leftIndent === undefined) {
    // Without explicit setIndents, derive from list configuration.
    // Without setLevelIndents, model the "much deeper" default indent Kevin saw.
    const indents = ref.list.levelIndents.get(ref.level) ?? { textIndent: 90 + 36 * ref.level, bulletIndent: 72 + 36 * ref.level };
    leftIndent = indents.textIndent;
    firstLineIndent = indents.bulletIndent - indents.textIndent;
  }
  return {
    text: paragraphText(p).replace(/\n/g, ''),
    isListItem: Boolean(ref),
    styleBuiltIn: p.styleBuiltIn ?? (p.styleName ? 'Other' : 'Normal'),
    leftIndent,
    firstLineIndent,
    listId: ref ? ref.list.id : null,
    level: ref ? ref.level : null,
    listString: ref ? ref.list.listString(p) : null
  };
}
