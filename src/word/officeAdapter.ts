/// <reference types="office-js" />
import type {
  CursorInfo,
  ListState,
  ParagraphIntent,
  ReadBackParagraph,
  ReadBackResult,
  ReadBackTable,
  RunFormat,
  TableStyleSettings,
  WordDocument,
  WordListHandle,
  WordParagraphHandle,
  WordTableHandle
} from './adapter';
import type { Block } from '../model';
import { describeError, insertBlocks, type InsertOptions, type InsertReport } from './insert';
import { formatLook, lookMismatches, parseTblLook, type TableLook } from './tableLook';

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

export class OfficeParagraphHandle implements WordParagraphHandle {
  // Not private: OfficeWordDocument needs the raw Word.Paragraph proxy back
  // for reads (list state, verify/repair, read-back).
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
    const withBreaks = toWordText(text);
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

  attachToList(listId: number, level: number): void {
    this.paragraph.attachToList(listId, level);
  }

  setListLevel(level: number): void {
    // The throwing variant: only called once a prior sync has made this
    // paragraph a list item.
    this.paragraph.listItem.level = level;
  }

  detachFromList(): void {
    this.paragraph.detachFromList();
  }

  setIndents(leftIndentPt: number, firstLineIndentPt: number): void {
    this.paragraph.leftIndent = leftIndentPt;
    this.paragraph.firstLineIndent = firstLineIndentPt;
  }
}

class OfficeListHandle implements WordListHandle {
  constructor(
    private readonly list: Word.List,
    private readonly context: Word.RequestContext
  ) {}

  async getId(): Promise<number> {
    this.list.load('id');
    await this.context.sync();
    return this.list.id;
  }

  setLevelBullet(level: number): void {
    this.list.setLevelBullet(level, Word.ListBullet.solid);
  }

  setLevelNumbering(level: number): void {
    this.list.setLevelNumbering(level, Word.ListNumbering.arabic, [level, '.']);
  }

  setLevelIndents(level: number, textIndentPt: number, bulletIndentPt: number): void {
    this.list.setLevelIndents(level, textIndentPt, bulletIndentPt);
  }

  setLevelStartingNumber(level: number, startingNumber: number): void {
    this.list.setLevelStartingNumber(level, startingNumber);
  }
}

class OfficeTableHandle implements WordTableHandle {
  constructor(
    private readonly table: Word.Table,
    private readonly context: Word.RequestContext
  ) {}

  async applyStyle(settings: TableStyleSettings): Promise<string[]> {
    const notes: string[] = [];

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
      notes.push(`table look: could not set flags (${String(err).substring(0, 50)})`);
      return notes;
    }

    // Verify and correct the table look by reading back the OOXML and comparing
    // with what we intended to set.
    await this.verifyAndCorrectTableLook(settings, notes);

    return notes;
  }

  private async verifyAndCorrectTableLook(settings: TableStyleSettings, notes: string[]): Promise<void> {
    // Build the target: what we intended to set
    const target: TableLook = {
      firstRow: true,
      lastRow: settings.styleTotalRow,
      firstColumn: settings.styleFirstColumn,
      lastColumn: settings.styleLastColumn,
      noHBand: !settings.styleBandedRows,
      noVBand: !settings.styleBandedColumns
    };

    // Read the actual look from the OOXML
    let actual = await this.readActualTableLook();
    if (!actual) {
      notes.push('table look: could not read tblLook');
      return;
    }

    notes.push(`table look (as set): ${formatLook(actual)}`);

    // Check for mismatches
    const mismatches = lookMismatches(actual, target);
    if (!mismatches.length) {
      return; // Perfect match
    }

    notes.push(`table look: mismatches in [${mismatches.join(', ')}]`);

    // Try to correct mismatches in firstColumn and lastColumn
    if (mismatches.includes('firstColumn') || mismatches.includes('lastColumn')) {
      await this.trySwapColumns(target, notes);
      // Re-read after correction attempt
      actual = await this.readActualTableLook();
      if (actual) {
        notes.push(`table look (after column swap): ${formatLook(actual)}`);
        // Update target and check again
        const newMismatches = lookMismatches(actual, target);
        if (!newMismatches.length) {
          return; // Correction successful
        }
      }
    }

    // Try to correct mismatches in noHBand and noVBand
    if (mismatches.includes('noHBand') || mismatches.includes('noVBand')) {
      await this.trySwapBands(target, notes);
      // Re-read after correction attempt
      actual = await this.readActualTableLook();
      if (actual) {
        notes.push(`table look (after band swap): ${formatLook(actual)}`);
      }
    }
  }

  private async readActualTableLook(): Promise<TableLook | null> {
    try {
      const range = this.table.getRange();
      const ooxml = range.getOoxml();
      await this.context.sync();
      return parseTblLook(ooxml.value);
    } catch (err) {
      console.warn('AI Paste: could not read table OOXML to verify table look.', err);
      return null;
    }
  }

  private async trySwapColumns(target: TableLook, notes: string[]): Promise<void> {
    try {
      // Swap firstColumn and lastColumn. Use the intended values from
      // `target`: the proxy's own properties aren't loaded, so reading them
      // would throw PropertyNotLoaded.
      this.table.styleFirstColumn = target.lastColumn;
      this.table.styleLastColumn = target.firstColumn;
      await this.context.sync();

      // Read back to check if the swap fixed it
      const newActual = await this.readActualTableLook();
      if (newActual) {
        const target2: TableLook = { ...target };
        const newMismatches = lookMismatches(newActual, target2);
        if (!newMismatches.includes('firstColumn') && !newMismatches.includes('lastColumn')) {
          // Swap was successful
          notes.push(`table look: first/last column were swapped by Word; corrected by swapping the flags`);
          return;
        }
      }

      // Swap did not fix it; restore original values
      this.table.styleFirstColumn = target.firstColumn;
      this.table.styleLastColumn = target.lastColumn;
      await this.context.sync();
      notes.push(`table look: could not correct first/last column (after swap: ${newActual ? formatLook(newActual) : 'unreadable'})`);
    } catch (err) {
      console.warn('AI Paste: error trying to swap table column flags.', err);
      notes.push(`table look: column swap attempt failed (${String(err).substring(0, 80)})`);
    }
  }

  private async trySwapBands(target: TableLook, notes: string[]): Promise<void> {
    try {
      // Swap styleBandedRows and styleBandedColumns, again from the intended
      // values (noHBand = !bandedRows, noVBand = !bandedColumns).
      this.table.styleBandedRows = !target.noVBand;
      this.table.styleBandedColumns = !target.noHBand;
      await this.context.sync();

      // Read back to check if the swap fixed it
      const newActual = await this.readActualTableLook();
      if (newActual) {
        const newMismatches = lookMismatches(newActual, target);
        if (!newMismatches.includes('noHBand') && !newMismatches.includes('noVBand')) {
          // Swap was successful
          notes.push(`table look: banding was reversed by Word; corrected by swapping the flags`);
          return;
        }
      }

      // Swap did not fix it; restore original values
      this.table.styleBandedRows = !target.noHBand;
      this.table.styleBandedColumns = !target.noVBand;
      await this.context.sync();
      notes.push(`table look: could not correct banding (after swap: ${newActual ? formatLook(newActual) : 'unreadable'})`);
    } catch (err) {
      console.warn('AI Paste: error trying to swap table band flags.', err);
      notes.push(`table look: band swap attempt failed (${String(err).substring(0, 80)})`);
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

function raw(handle: WordParagraphHandle): Word.Paragraph {
  return (handle as OfficeParagraphHandle).paragraph;
}

function numberOrNull(n: unknown): number | null {
  return typeof n === 'number' && Number.isFinite(n) ? n : null;
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
    paragraph.load('text,isListItem');
    await this.context.sync();

    if (paragraph.isNullObject) {
      const fallback = this.context.document.body.paragraphs.getLastOrNullObject();
      fallback.load('text,isListItem');
      await this.context.sync();
      paragraph = fallback;
    }

    const isEmpty = !paragraph.isNullObject && paragraph.text.trim().length === 0;
    const isListItem = !paragraph.isNullObject && paragraph.isListItem;
    return { paragraph: new OfficeParagraphHandle(paragraph, this.context), isEmpty, isListItem };
  }

  async commit(): Promise<void> {
    await this.context.sync();
  }

  async readListState(paragraphs: WordParagraphHandle[]): Promise<ListState[]> {
    const raws = paragraphs.map(raw);
    const lists = raws.map((p) => p.listOrNullObject);
    const items = raws.map((p) => p.listItemOrNullObject);
    raws.forEach((p) => p.load('isListItem'));
    lists.forEach((l) => l.load('id'));
    items.forEach((i) => i.load('level'));
    await this.context.sync();
    return raws.map((p, i) => ({
      isListItem: p.isListItem,
      listId: lists[i].isNullObject ? null : lists[i].id,
      level: items[i].isNullObject ? null : items[i].level
    }));
  }

  async verifyAndRepairListItems(records: ParagraphIntent[]): Promise<number> {
    if (!records.length) return 0;

    const withRaw = records.map((r) => ({ ...r, raw: raw(r.handle) }));
    withRaw.forEach((r) => r.raw.load('isListItem'));
    await this.context.sync();

    const toRepair = withRaw.filter((r) => !r.intendedListItem && r.raw.isListItem);
    for (const r of toRepair) {
      r.handle.detachFromList();
      r.reapplyStyle?.();
    }
    if (toRepair.length) {
      await this.context.sync();
    }
    return toRepair.length;
  }

  async readBack(first: WordParagraphHandle, last: WordParagraphHandle): Promise<ReadBackResult> {
    const notes: string[] = [];
    const range = raw(first)
      .getRange(Word.RangeLocation.whole)
      .expandTo(raw(last).getRange(Word.RangeLocation.whole));

    const collection = range.paragraphs;
    collection.load('items/text,items/isListItem,items/styleBuiltIn,items/tableNestingLevel,items/leftIndent,items/firstLineIndent');
    await this.context.sync();

    const body = collection.items.filter((p) => !(p.tableNestingLevel > 0));

    // List details only for list items, in a separate sync so a failure
    // here still leaves the basic read-back usable.
    const listItems = body.map((p) => (p.isListItem ? p : null));
    const ids = new Map<Word.Paragraph, Word.List>();
    const levels = new Map<Word.Paragraph, Word.ListItem>();
    try {
      for (const p of listItems) {
        if (!p) continue;
        const list = p.listOrNullObject;
        list.load('id');
        ids.set(p, list);
        const item = p.listItemOrNullObject;
        item.load('level,listString');
        levels.set(p, item);
      }
      if (ids.size) await this.context.sync();
    } catch (err) {
      notes.push(`list ids/levels not readable (${describeError(err)})`);
      ids.clear();
      levels.clear();
    }

    const paragraphs: ReadBackParagraph[] = body.map((p) => {
      const list = ids.get(p);
      const item = levels.get(p);
      return {
        text: p.text,
        isListItem: p.isListItem,
        styleBuiltIn: String(p.styleBuiltIn),
        leftIndent: numberOrNull(p.leftIndent),
        firstLineIndent: numberOrNull(p.firstLineIndent),
        listId: list && !list.isNullObject ? list.id : null,
        level: item && !item.isNullObject ? item.level : null,
        listString: item && !item.isNullObject ? item.listString : null
      };
    });

    let tables: ReadBackTable[] = [];
    try {
      const tableCollection = range.tables;
      tableCollection.load('items/rowCount,items/headerRowCount');
      await this.context.sync();
      tables = tableCollection.items.map((t) => ({ rowCount: t.rowCount, headerRowCount: t.headerRowCount }));
    } catch (err) {
      notes.push(`tables not readable (${describeError(err)})`);
    }

    return { paragraphs, tables, notes };
  }
}

/** Inserts the given blocks at the current cursor position in the active Word document. */
export async function insertBlocksInWord(blocks: Block[], options: InsertOptions): Promise<InsertReport> {
  return Word.run(async (context) => {
    const doc = new OfficeWordDocument(context);
    return insertBlocks(doc, blocks, options);
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
