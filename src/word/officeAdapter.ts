/// <reference types="office-js" />
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
} from './adapter';
import type { Block } from '../model';
import { insertBlocks, type InsertOptions } from './insert';
import { buildListOoxml } from './ooxml';

/** True once, cheap: is the Word host new enough for everything this add-in needs? */
export function isWordApi13Supported(): boolean {
  return Boolean(Office.context?.requirements?.isSetSupported('WordApi', '1.3'));
}

function isWordApi15Supported(): boolean {
  return Boolean(Office.context?.requirements?.isSetSupported('WordApi', '1.5'));
}

function applyFormat(range: Word.Range, format: RunFormat): void {
  range.font.bold = format.bold;
  range.font.italic = format.italic;
  if (format.fontName) range.font.name = format.fontName;
  if (format.link) range.hyperlink = format.link;
}

class OfficeParagraphHandle implements WordParagraphHandle {
  // Not private: OfficeWordDocument needs the raw Word.Paragraph proxy back
  // for verifyAndRepairListItems (load/read isListItem) and insertOoxmlList
  // (as an insertion anchor).
  constructor(
    readonly paragraph: Word.Paragraph,
    private readonly context: Word.RequestContext
  ) {}

  setStyleBuiltIn(name: string): void {
    this.paragraph.styleBuiltIn = name as Word.BuiltInStyleName;
  }

  setStyleByName(name: string): void {
    this.paragraph.style = name;
  }

  insertRun(text: string, format: RunFormat): void {
    // Word represents a manual line break (Shift+Enter) within a paragraph's
    // text as a vertical tab (U+000B), both when reading Range.text back and
    // when inserting it via insertText.
    const withBreaks = text.replace(/\n/g, '\v');
    if (!withBreaks) return;
    const range = this.paragraph.insertText(withBreaks, Word.InsertLocation.end);
    applyFormat(range, format);
  }

  setWholeTextFormat(format: RunFormat): void {
    applyFormat(this.paragraph.getRange(), format);
  }

  insertParagraphAfter(text: string): WordParagraphHandle {
    return new OfficeParagraphHandle(this.paragraph.insertParagraph(toWordText(text), Word.InsertLocation.after), this.context);
  }

  insertTableAfter(values: string[][]): WordTableHandle {
    const rowCount = values.length;
    const colCount = values[0]?.length ?? 0;
    const table = this.paragraph.insertTable(rowCount, colCount, Word.InsertLocation.after, values);
    return new OfficeTableHandle(table, this.context);
  }

  startNewList(): WordListHandle {
    return new OfficeListHandle(this.paragraph.startNewList(), this.context);
  }

  setListLevel(level: number): void {
    // The throwing variant: only valid (and only called) once a prior sync
    // has confirmed this paragraph really is a list item.
    this.paragraph.listItem.level = level;
  }

  detachFromList(): void {
    this.paragraph.detachFromList();
  }
}

class OfficeListHandle implements WordListHandle {
  private readonly configuredLevels = new Set<number>();

  constructor(
    private readonly list: Word.List,
    private readonly context: Word.RequestContext
  ) {}

  insertItemAfter(text: string): WordParagraphHandle {
    return new OfficeParagraphHandle(this.list.insertParagraph(toWordText(text), Word.InsertLocation.end), this.context);
  }

  insertParagraphAfter(text: string): WordParagraphHandle {
    // "After" on a List (like on a Table) inserts relative to the whole
    // list's boundary, not as a new member — unlike "Start"/"End", which
    // insert inside it. This is what lets a plain body paragraph follow a
    // list without becoming another bullet/number itself.
    return new OfficeParagraphHandle(this.list.insertParagraph(toWordText(text), Word.InsertLocation.after), this.context);
  }

  async ensureBulletLevel(level: number): Promise<void> {
    if (this.configuredLevels.has(level)) return;
    this.configuredLevels.add(level);
    try {
      this.list.setLevelBullet(level, Word.ListBullet.solid);
      await this.context.sync();
    } catch (err) {
      console.warn(`AI Paste: setLevelBullet(${level}) failed; leaving Word's default bullet formatting.`, err);
    }
  }

  async ensureNumberLevel(level: number, startAt?: number): Promise<void> {
    if (this.configuredLevels.has(level)) return;
    this.configuredLevels.add(level);
    try {
      this.list.setLevelNumbering(level, Word.ListNumbering.arabic, [level, '.']);
      if (startAt && startAt !== 1) this.list.setLevelStartingNumber(level, startAt);
      await this.context.sync();
    } catch (err) {
      console.warn(`AI Paste: setLevelNumbering(${level}) failed; leaving Word's default number formatting.`, err);
    }
  }
}

class OfficeTableHandle implements WordTableHandle {
  constructor(
    private readonly table: Word.Table,
    private readonly context: Word.RequestContext
  ) {}

  async applyStyle(settings: TableStyleSettings): Promise<void> {
    // The table must exist server-side before it can be styled.
    await this.context.sync();

    try {
      if (settings.styleBuiltIn) {
        this.table.styleBuiltIn = settings.styleBuiltIn as Word.BuiltInStyleName;
      } else if (settings.customName) {
        this.table.style = settings.customName;
      }
      await this.context.sync();
    } catch (err) {
      console.warn(`AI Paste: styleBuiltIn "${settings.styleBuiltIn}" failed on the table; trying its display name.`, err);
      try {
        if (settings.styleBuiltIn) this.table.style = toDisplayName(settings.styleBuiltIn);
        await this.context.sync();
      } catch (err2) {
        console.warn('AI Paste: table style fallback also failed; leaving Word\'s default table look.', err2);
      }
    }

    try {
      this.table.headerRowCount = settings.headerRowCount;
      this.table.styleFirstColumn = settings.styleFirstColumn;
      this.table.styleBandedRows = settings.styleBandedRows;
      this.table.styleBandedColumns = settings.styleBandedColumns;
      this.table.styleLastColumn = settings.styleLastColumn;
      this.table.styleTotalRow = settings.styleTotalRow;
      await this.context.sync();
    } catch (err) {
      console.warn('AI Paste: table style flags (header row / banding) failed to apply.', err);
    }
  }

  insertParagraphAfter(text: string): WordParagraphHandle {
    return new OfficeParagraphHandle(this.table.insertParagraph(toWordText(text), Word.InsertLocation.after), this.context);
  }
}

/** "GridTable4" -> "Grid Table 4", best-effort last-resort fallback for an English-UI Word only. */
function toDisplayName(builtInStyleName: string): string {
  return builtInStyleName.replace(/([a-z])([A-Z0-9])/g, '$1 $2').replace(/(\d)/g, ' $1');
}

function toWordText(text: string): string {
  return text.replace(/\n/g, '\v');
}

export class OfficeWordDocument implements WordDocument {
  constructor(private readonly context: Word.RequestContext) {}

  async getCursor(): Promise<CursorInfo> {
    const selection = this.context.document.getSelection();
    selection.load('isEmpty');
    await this.context.sync();

    // Empty-string insertText is a known GeneralException source on some
    // Word hosts — never call it. Only delete a genuinely non-empty
    // selection; a collapsed cursor is left exactly as it is.
    if (!selection.isEmpty) {
      selection.delete();
      await this.context.sync();
    }

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
        ? new OfficeListHandle(paragraph.listOrNullObject, this.context)
        : null;
    return { paragraph: new OfficeParagraphHandle(paragraph, this.context), isEmpty, list };
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

  async insertOoxmlList(
    anchor: WordParagraphHandle,
    anchorIsEmpty: boolean,
    items: OoxmlListItemInput[]
  ): Promise<{ items: WordParagraphHandle[]; list: WordListHandle | null }> {
    const xml = buildListOoxml(items);
    const anchorRaw = (anchor as OfficeParagraphHandle).paragraph;
    const insertedRange = anchorIsEmpty
      ? anchorRaw.insertOoxml(xml, Word.InsertLocation.replace)
      : anchorRaw.getRange().insertOoxml(xml, Word.InsertLocation.after);

    const paragraphs = insertedRange.paragraphs;
    paragraphs.load('items/text,items/listOrNullObject/isNullObject');
    await this.context.sync();

    const itemCount = items.length;
    const itemParagraphs = paragraphs.items.slice(0, itemCount);
    for (const p of itemParagraphs) {
      p.styleBuiltIn = 'ListParagraph' as Word.BuiltInStyleName;
    }

    // insertOoxml can leave one stray empty paragraph after the inserted
    // content; remove it if present (best-effort — never blocks this path).
    if (paragraphs.items.length > itemCount) {
      const extra = paragraphs.items[itemCount];
      if (extra.text.trim() === '') {
        try {
          extra.delete();
        } catch (err) {
          console.warn('AI Paste: could not remove the stray paragraph after an OOXML list insert.', err);
        }
      }
    }
    await this.context.sync();

    const last = itemParagraphs[itemParagraphs.length - 1];
    const list = last && !last.listOrNullObject.isNullObject ? new OfficeListHandle(last.listOrNullObject, this.context) : null;

    return { items: itemParagraphs.map((p) => new OfficeParagraphHandle(p, this.context)), list };
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
