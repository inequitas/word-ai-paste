import type {
  CursorInfo,
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

class FakeParagraphHandle implements WordParagraphHandle {
  constructor(private readonly model: FakeDocumentModel, readonly para: FakeParagraph) {}

  setStyleBuiltIn(name: string): void {
    this.para.styleBuiltIn = name;
    this.para.styleName = undefined;
  }

  setStyleByName(name: string): void {
    this.para.styleName = name;
    this.para.styleBuiltIn = undefined;
  }

  insertRun(text: string, format: RunFormat): void {
    const segments = text.split('\n');
    segments.forEach((seg, i) => {
      if (i > 0) this.para.lines.push([]);
      if (seg) {
        this.para.lines[this.para.lines.length - 1].push({
          text: seg,
          bold: format.bold,
          italic: format.italic,
          fontName: format.fontName,
          link: format.link
        });
      }
    });
  }

  insertParagraphAfter(): WordParagraphHandle {
    const p = emptyParagraph();
    insertAfter(this.model, this.para, p);
    return new FakeParagraphHandle(this.model, p);
  }

  insertTableAfter(values: string[][]): WordTableHandle {
    const t: FakeTable = { kind: 'table', values, style: null };
    insertAfter(this.model, this.para, t);
    return new FakeTableHandle(this.model, t);
  }

  startNewList(): WordListHandle {
    const list = new FakeList(this.model);
    this.para.listRef = { list, level: 0 };
    list.members.push(this.para);
    return list;
  }

  setListLevel(level: number): void {
    if (this.para.listRef) this.para.listRef.level = level;
  }
}

class FakeList implements WordListHandle {
  members: FakeParagraph[] = [];
  readonly bulletLevels = new Set<number>();
  readonly numberLevels = new Map<number, number>();

  constructor(private readonly model: FakeDocumentModel) {}

  insertItemAfter(): WordParagraphHandle {
    const last = this.members[this.members.length - 1];
    const p = emptyParagraph();
    p.listRef = { list: this, level: 0 };
    insertAfter(this.model, last, p);
    this.members.push(p);
    return new FakeParagraphHandle(this.model, p);
  }

  ensureBulletLevel(level: number): void {
    this.bulletLevels.add(level);
  }

  ensureNumberLevel(level: number, startAt?: number): void {
    if (this.numberLevels.has(level)) return;
    this.numberLevels.set(level, startAt ?? 1);
  }
}

class FakeTableHandle implements WordTableHandle {
  constructor(private readonly model: FakeDocumentModel, readonly table: FakeTable) {}

  applyStyle(settings: TableStyleSettings): void {
    this.table.style = settings;
  }

  insertParagraphAfter(): WordParagraphHandle {
    const p = emptyParagraph();
    insertAfter(this.model, this.table, p);
    return new FakeParagraphHandle(this.model, p);
  }
}

export class FakeWordDocument implements WordDocument {
  readonly model = new FakeDocumentModel();
  private readonly anchor: FakeParagraph;
  commitCount = 0;

  constructor(options: { anchorText?: string } = {}) {
    this.anchor = emptyParagraph();
    if (options.anchorText) {
      this.anchor.lines[0].push({ text: options.anchorText, bold: false, italic: false });
    }
    this.model.blocks.push(this.anchor);
  }

  async getCursor(): Promise<CursorInfo> {
    const isEmpty = this.anchor.lines.length === 1 && this.anchor.lines[0].length === 0;
    return { paragraph: new FakeParagraphHandle(this.model, this.anchor), isEmpty };
  }

  async commit(): Promise<void> {
    this.commitCount++;
  }
}
