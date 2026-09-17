import type {
  CursorInfo,
  OoxmlListItemInput,
  ParagraphIntent,
  RunFormat,
  TableStyleSettings,
  WordDocument,
  WordListHandle,
  WordParagraphHandle,
  WordTableHandle
} from '../src/word/adapter';

/**
 * A tiny in-memory recording implementation of the Word adapter interfaces,
 * used to unit-test src/word/insert.ts without a real Word runtime.
 *
 * It deliberately reproduces one real-Word quirk: calling
 * `insertParagraphAfter()` on a paragraph that is currently a list item
 * creates *another list item* of the same list/level, the same way
 * pressing Enter at the end of a list line does in real Word. That's the
 * exact behavior src/word/insert.ts has to route around (via
 * `WordListHandle.insertParagraphAfter()` instead) and its
 * verify-and-repair pass has to catch as a safety net.
 *
 * It can also simulate the GeneralException Kevin actually hit: call
 * `FakeWordDocument.failNext('startNewList' | 'insertItemAfter' |
 * 'setListLevel')` to make the next matching call throw a fake
 * Office.js-shaped error, so tests can exercise insert.ts's OOXML fallback
 * path without a real Word runtime.
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
  /** Set on paragraphs created by the OOXML fallback, for assertions. */
  viaOoxml?: boolean;
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
}

function insertAfter(model: FakeDocumentModel, anchor: FakeBlock, block: FakeBlock): void {
  const idx = model.blocks.indexOf(anchor);
  if (idx === -1) throw new Error('fake adapter: anchor block not found');
  model.blocks.splice(idx + 1, 0, block);
}

function emptyParagraph(): FakeParagraph {
  return { kind: 'paragraph', lines: [[]] };
}

function fillFirstRun(p: FakeParagraph, text: string, format: RunFormat): void {
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

class FakeParagraphHandle implements WordParagraphHandle {
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
    fillFirstRun(this.para, text, format);
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
    const p = emptyParagraph();
    if (text) fillFirstRun(p, text, { bold: false, italic: false });
    // Mirrors real Word: a paragraph inserted right after a list item
    // becomes another item of that same list/level unless something
    // explicitly detaches it.
    if (this.para.listRef) {
      const { list, level } = this.para.listRef;
      p.listRef = { list, level };
      list.members.push(p);
    }
    insertAfter(this.model, this.para, p);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }

  insertTableAfter(values: string[][]): WordTableHandle {
    const t: FakeTable = { kind: 'table', values, style: null };
    insertAfter(this.model, this.para, t);
    return new FakeTableHandle(this.model, this.doc, t);
  }

  startNewList(): WordListHandle {
    if (this.doc.consumeFailure('startNewList')) throw new FakeOfficeError('Word.Paragraph.startNewList');
    const list = new FakeList(this.model, this.doc);
    this.para.listRef = { list, level: 0 };
    list.members.push(this.para);
    return list;
  }

  setListLevel(level: number): void {
    if (this.doc.consumeFailure('setListLevel')) throw new FakeOfficeError('Word.ListItem.level');
    if (this.para.listRef) this.para.listRef.level = level;
  }

  detachFromList(): void {
    if (!this.para.listRef) return;
    const { list } = this.para.listRef;
    const idx = list.members.indexOf(this.para);
    if (idx !== -1) list.members.splice(idx, 1);
    this.para.listRef = undefined;
  }
}

class FakeList implements WordListHandle {
  members: FakeParagraph[] = [];
  readonly bulletLevels = new Set<number>();
  readonly numberLevels = new Map<number, number>();

  constructor(
    private readonly model: FakeDocumentModel,
    private readonly doc: FakeWordDocument
  ) {}

  insertItemAfter(text: string): WordParagraphHandle {
    if (this.doc.consumeFailure('insertItemAfter')) throw new FakeOfficeError('Word.List.insertParagraph');
    const last = this.members[this.members.length - 1];
    const p = emptyParagraph();
    if (text) fillFirstRun(p, text, { bold: false, italic: false });
    p.listRef = { list: this, level: 0 };
    insertAfter(this.model, last, p);
    this.members.push(p);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }

  insertParagraphAfter(text: string): WordParagraphHandle {
    // Inserted after the whole list (after its last member), and — unlike
    // insertItemAfter — never a member of it.
    const last = this.members[this.members.length - 1];
    const p = emptyParagraph();
    if (text) fillFirstRun(p, text, { bold: false, italic: false });
    insertAfter(this.model, last, p);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }

  async ensureBulletLevel(level: number): Promise<void> {
    this.bulletLevels.add(level);
  }

  async ensureNumberLevel(level: number, startAt?: number): Promise<void> {
    if (this.numberLevels.has(level)) return;
    this.numberLevels.set(level, startAt ?? 1);
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
    const p = emptyParagraph();
    if (text) fillFirstRun(p, text, { bold: false, italic: false });
    insertAfter(this.model, this.table, p);
    return new FakeParagraphHandle(this.model, this.doc, p);
  }
}

type FailableAction = 'startNewList' | 'insertItemAfter' | 'setListLevel';

export class FakeWordDocument implements WordDocument {
  readonly model = new FakeDocumentModel();
  private readonly anchor: FakeParagraph;
  private readonly anchorList: FakeList | null = null;
  commitCount = 0;
  /** Diagnostics for tests: how many paragraphs the repair pass detached + restyled, and any warnings it logged. */
  repairedCount = 0;
  warnings: string[] = [];
  /** Records of every insertOoxmlList call, for assertions. */
  ooxmlCalls: { anchorIsEmpty: boolean; items: OoxmlListItemInput[] }[] = [];

  private failing = new Set<FailableAction>();

  constructor(options: { anchorText?: string; anchorInList?: boolean } = {}) {
    this.anchor = emptyParagraph();
    if (options.anchorText) {
      this.anchor.lines[0].push({ text: options.anchorText, bold: false, italic: false });
    }
    if (options.anchorInList) {
      const list = new FakeList(this.model, this);
      this.anchor.listRef = { list, level: 0 };
      list.members.push(this.anchor);
      this.anchorList = list;
    }
    this.model.blocks.push(this.anchor);
  }

  /** Test hook: make the next call to this action throw a FakeOfficeError (one-shot). */
  failNext(action: FailableAction): void {
    this.failing.add(action);
  }

  /** Internal: called by the fake handles; returns true (and clears the flag) exactly once per failNext() call. */
  consumeFailure(action: FailableAction): boolean {
    if (!this.failing.has(action)) return false;
    this.failing.delete(action);
    return true;
  }

  async getCursor(): Promise<CursorInfo> {
    const isEmpty = this.anchor.lines.length === 1 && this.anchor.lines[0].length === 0;
    return {
      paragraph: new FakeParagraphHandle(this.model, this, this.anchor),
      isEmpty,
      list: this.anchorList
    };
  }

  async commit(): Promise<void> {
    this.commitCount++;
  }

  async verifyAndRepairListItems(records: ParagraphIntent[]): Promise<void> {
    for (const r of records) {
      const para = (r.handle as FakeParagraphHandle).para;
      const isListItem = Boolean(para.listRef);
      if (!r.intendedListItem && isListItem) {
        r.handle.detachFromList();
        r.reapplyStyle?.();
        this.repairedCount++;
      } else if (r.intendedListItem && !isListItem) {
        this.warnings.push(`expected list item, found none: ${paragraphText(para)}`);
      }
    }
  }

  async insertOoxmlList(
    anchor: WordParagraphHandle,
    anchorIsEmpty: boolean,
    items: OoxmlListItemInput[]
  ): Promise<{ items: WordParagraphHandle[]; list: WordListHandle | null }> {
    this.ooxmlCalls.push({ anchorIsEmpty, items });

    const anchorPara = (anchor as FakeParagraphHandle).para;
    const list = new FakeList(this.model, this);
    const created: FakeParagraph[] = [];

    const fillFromItem = (p: FakeParagraph, item: OoxmlListItemInput): void => {
      p.styleBuiltIn = 'ListParagraph';
      p.styleName = undefined;
      p.viaOoxml = true;
      const first = item.inlines[0];
      const text = item.inlines.map((i) => i.text).join('');
      if (text) fillFirstRun(p, text, { bold: !!first?.bold, italic: !!first?.italic });
      p.listRef = { list, level: item.level };
      list.members.push(p);
    };

    let anchorBlock: FakeBlock;
    let startFromIndex: number;

    if (anchorIsEmpty) {
      // "Replace": reuse the anchor paragraph itself as item 0.
      anchorPara.lines = [[]];
      if (items[0]) fillFromItem(anchorPara, items[0]);
      created.push(anchorPara);
      anchorBlock = anchorPara;
      startFromIndex = 1;
    } else {
      anchorBlock = anchorPara;
      startFromIndex = 0;
    }

    for (let i = startFromIndex; i < items.length; i++) {
      const p = emptyParagraph();
      fillFromItem(p, items[i]);
      insertAfter(this.model, anchorBlock, p);
      anchorBlock = p;
      created.push(p);
    }

    return { items: created.map((p) => new FakeParagraphHandle(this.model, this, p)), list };
  }
}
