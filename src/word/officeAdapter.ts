/// <reference types="office-js" />
import type {
  CursorInfo,
  ParagraphIntent,
  RunFormat,
  TableStyleSettings,
  WordDocument,
  WordListHandle,
  WordParagraphHandle,
  WordTableHandle
} from './adapter';
import type { Block } from '../model';
import { insertBlocks, type InsertOptions } from './insert';

/** True once, cheap: is the Word host new enough for everything this add-in needs? */
export function isWordApi13Supported(): boolean {
  return Boolean(Office.context?.requirements?.isSetSupported('WordApi', '1.3'));
}

function isWordApi15Supported(): boolean {
  return Boolean(Office.context?.requirements?.isSetSupported('WordApi', '1.5'));
}

class OfficeParagraphHandle implements WordParagraphHandle {
  // Not private: OfficeWordDocument.verifyAndRepairListItems needs the raw
  // Word.Paragraph proxy back to load()/read isListItem on it.
  constructor(readonly paragraph: Word.Paragraph) {}

  setStyleBuiltIn(name: string): void {
    this.paragraph.styleBuiltIn = name as Word.BuiltInStyleName;
  }

  setStyleByName(name: string): void {
    this.paragraph.style = name;
  }

  insertRun(text: string, format: RunFormat): void {
    // Word represents a manual line break (Shift+Enter) within a paragraph's
    // text as a vertical tab (U+000B), both when reading Range.text back and
    // when inserting it via insertText — there is no "insert a break at the
    // end of this range" overload (Paragraph/Range.insertBreak only accept
    // Before/After, which splice in a break relative to the whole
    // paragraph/range, not inside it).
    const withBreaks = text.replace(/\n/g, '\v');
    if (!withBreaks) return;
    const range = this.paragraph.insertText(withBreaks, Word.InsertLocation.end);
    range.font.bold = format.bold;
    range.font.italic = format.italic;
    if (format.fontName) range.font.name = format.fontName;
    if (format.link) range.hyperlink = format.link;
  }

  insertParagraphAfter(): WordParagraphHandle {
    return new OfficeParagraphHandle(this.paragraph.insertParagraph('', Word.InsertLocation.after));
  }

  insertTableAfter(values: string[][]): WordTableHandle {
    const rowCount = values.length;
    const colCount = values[0]?.length ?? 0;
    const table = this.paragraph.insertTable(rowCount, colCount, Word.InsertLocation.after, values);
    return new OfficeTableHandle(table);
  }

  startNewList(): WordListHandle {
    return new OfficeListHandle(this.paragraph.startNewList());
  }

  setListLevel(level: number): void {
    this.paragraph.listItemOrNullObject.level = level;
  }

  detachFromList(): void {
    this.paragraph.detachFromList();
  }
}

class OfficeListHandle implements WordListHandle {
  private readonly configuredLevels = new Set<number>();

  constructor(private readonly list: Word.List) {}

  insertItemAfter(): WordParagraphHandle {
    return new OfficeParagraphHandle(this.list.insertParagraph('', Word.InsertLocation.end));
  }

  insertParagraphAfter(): WordParagraphHandle {
    // "After" on a List (like on a Table) inserts relative to the whole
    // list's boundary, not as a new member — unlike "Start"/"End", which
    // insert inside it. This is what lets a plain body paragraph follow a
    // list without becoming another bullet/number itself.
    return new OfficeParagraphHandle(this.list.insertParagraph('', Word.InsertLocation.after));
  }

  ensureBulletLevel(level: number): void {
    const key = level;
    if (this.configuredLevels.has(key)) return;
    this.configuredLevels.add(key);
    this.list.setLevelBullet(level, Word.ListBullet.solid);
  }

  ensureNumberLevel(level: number, startAt?: number): void {
    if (this.configuredLevels.has(level)) return;
    this.configuredLevels.add(level);
    this.list.setLevelNumbering(level, Word.ListNumbering.arabic, [level, '.']);
    if (startAt && startAt !== 1) this.list.setLevelStartingNumber(level, startAt);
  }
}

class OfficeTableHandle implements WordTableHandle {
  constructor(private readonly table: Word.Table) {}

  applyStyle(settings: TableStyleSettings): void {
    if (settings.styleBuiltIn) {
      this.table.styleBuiltIn = settings.styleBuiltIn as Word.BuiltInStyleName;
    } else if (settings.customName) {
      this.table.style = settings.customName;
    }
    this.table.headerRowCount = settings.headerRowCount;
    this.table.styleFirstColumn = settings.styleFirstColumn;
    this.table.styleBandedRows = settings.styleBandedRows;
    this.table.styleBandedColumns = settings.styleBandedColumns;
    this.table.styleLastColumn = settings.styleLastColumn;
    this.table.styleTotalRow = settings.styleTotalRow;
  }

  insertParagraphAfter(): WordParagraphHandle {
    return new OfficeParagraphHandle(this.table.insertParagraph('', Word.InsertLocation.after));
  }
}

export class OfficeWordDocument implements WordDocument {
  constructor(private readonly context: Word.RequestContext) {}

  async getCursor(): Promise<CursorInfo> {
    // Replacing a selection with "" is a no-op when the selection is
    // already collapsed (the ordinary paste-at-cursor case), and clears it
    // otherwise — the brief calls for replacing a non-empty selection
    // rather than inserting after it, and this does that without needing
    // a separate load+sync just to check whether there was one.
    this.context.document.getSelection().insertText('', Word.InsertLocation.replace);

    let paragraph = this.context.document.getSelection().paragraphs.getFirstOrNullObject();
    paragraph.load('text,listOrNullObject/isNullObject');
    await this.context.sync();

    if (paragraph.isNullObject) {
      const fallback = this.context.document.body.paragraphs.getLastOrNullObject();
      fallback.load('text,listOrNullObject/isNullObject');
      await this.context.sync();
      paragraph = fallback;
    }

    const isEmpty = !paragraph.isNullObject && paragraph.text.trim().length === 0;
    const list =
      !paragraph.isNullObject && !paragraph.listOrNullObject.isNullObject
        ? new OfficeListHandle(paragraph.listOrNullObject)
        : null;
    return { paragraph: new OfficeParagraphHandle(paragraph), isEmpty, list };
  }

  async commit(): Promise<void> {
    await this.context.sync();
  }

  async verifyAndRepairListItems(records: ParagraphIntent[]): Promise<void> {
    if (!records.length) return;

    const withRaw = records.map((r) => ({ ...r, raw: (r.handle as OfficeParagraphHandle).paragraph }));
    withRaw.forEach((r) => r.raw.load('isListItem,text'));
    await this.context.sync();

    const toRepair = withRaw.filter((r) => !r.intendedListItem && r.raw.isListItem);
    for (const r of toRepair) {
      r.handle.detachFromList();
      r.reapplyStyle?.();
    }
    if (toRepair.length) {
      await this.context.sync();
    }

    for (const r of withRaw) {
      if (r.intendedListItem && !r.raw.isListItem) {
        console.warn('AI Paste: expected paragraph to be a list item after insert, but it is not:', r.raw.text);
      }
    }
  }
}

/** Inserts the given blocks at the current cursor position in the active Word document. */
export async function insertBlocksInWord(blocks: Block[], options: InsertOptions): Promise<void> {
  await Word.run(async (context) => {
    const doc = new OfficeWordDocument(context);
    await insertBlocks(doc, blocks, options);
  });
}

/** Populates the body-style / table-style pickers from the current document. Best-effort: returns [] pre-WordApi 1.5. */
export async function loadDocumentStyleNames(): Promise<{ paragraphStyles: string[]; tableStyles: string[] }> {
  if (!isWordApi15Supported()) return { paragraphStyles: [], tableStyles: [] };

  return Word.run(async (context) => {
    const styles = context.document.getStyles();
    styles.load('items/nameLocal,items/type,items/builtIn');
    await context.sync();

    const skipParagraph = new Set(['Normal', 'Quote', 'List Paragraph', 'ListParagraph']);
    for (let n = 1; n <= 9; n++) skipParagraph.add(`Heading ${n}`);

    const paragraphStyles = new Set<string>();
    const tableStyles = new Set<string>();
    for (const style of styles.items) {
      if (style.type === Word.StyleType.paragraph && !skipParagraph.has(style.nameLocal)) {
        paragraphStyles.add(style.nameLocal);
      } else if (style.type === Word.StyleType.table) {
        tableStyles.add(style.nameLocal);
      }
    }
    return {
      paragraphStyles: Array.from(paragraphStyles).sort(),
      tableStyles: Array.from(tableStyles).sort()
    };
  });
}
