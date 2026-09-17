import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { insertBlocks, computeOrdinals, DEFAULT_INSERT_OPTIONS, InsertError, formatStep } from '../src/word/insert';
import { blockPlainText, type Block, type ListItemBlock } from '../src/model';
import { parseMarkdown } from '../src/parse/markdown';
import { normalize, DEFAULT_NORMALIZE_OPTIONS } from '../src/transform/normalize';
import type { WordParagraphHandle } from '../src/word/adapter';
import {
  FakeWordDocument,
  isParagraph,
  isTable,
  paragraphText,
  readBackOf,
  type FakeList,
  type FakeParagraph,
  type FakeParagraphHandle
} from './fakeWordAdapter';

const bullet = (text: string, listIndex = 1, level = 0): ListItemBlock => ({
  type: 'listItem',
  ordered: false,
  level,
  listIndex,
  inlines: [{ text }]
});
const numbered = (text: string, listIndex = 1, level = 0, start?: number): ListItemBlock => ({
  type: 'listItem',
  ordered: true,
  level,
  listIndex,
  start,
  inlines: [{ text }]
});
const para = (text: string): Block => ({ type: 'paragraph', inlines: text ? [{ text }] : [] });

// insert.ts reports read-back mismatches and failed formatting calls via console.warn.
beforeEach(() => {
  vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(() => {
  vi.restoreAllMocks();
});

function listString(p: FakeParagraph): string | null {
  return readBackOf(p).listString;
}

describe('insertBlocks — basics', () => {
  it('reuses an empty cursor paragraph for the first block instead of creating a new one', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('Hello')], DEFAULT_INSERT_OPTIONS);

    expect(doc.model.blocks).toHaveLength(1);
    const [p] = doc.paragraphs;
    expect(paragraphText(p)).toBe('Hello');
    expect(p.styleBuiltIn).toBe('Normal');
  });

  it('inserts after a non-empty cursor paragraph instead of overwriting it', async () => {
    const doc = new FakeWordDocument({ anchorText: 'existing text' });
    await insertBlocks(doc, [para('New')], DEFAULT_INSERT_OPTIONS);
    expect(doc.texts()).toEqual(['existing text', 'New']);
  });

  it('inserts before existing content that follows the cursor', async () => {
    const doc = new FakeWordDocument({ anchorText: '', trailingText: 'later content' });
    await insertBlocks(doc, [para('One'), para('Two')], DEFAULT_INSERT_OPTIONS);
    expect(doc.texts()).toEqual(['One', 'Two', 'later content']);
  });

  it('sets styleBuiltIn to HeadingN, clamped at 9', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'heading', level: 2, inlines: [{ text: 'A' }] },
      { type: 'heading', level: 15, inlines: [{ text: 'B' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [a, b] = doc.paragraphs;
    expect(a.styleBuiltIn).toBe('Heading2');
    expect(b.styleBuiltIn).toBe('Heading9');
  });

  it('always sets bold/italic explicitly (true and false) on every run', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [{ type: 'paragraph', inlines: [{ text: 'bold ', bold: true }, { text: 'plain' }] }], DEFAULT_INSERT_OPTIONS);
    const runs = doc.paragraphs[0].lines[0];
    expect(runs[0]).toMatchObject({ text: 'bold ', bold: true, italic: false });
    expect(runs[1]).toMatchObject({ text: 'plain', bold: false, italic: false });
  });

  it('sets a hyperlink on a linked run and a monospace font on a code run', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      {
        type: 'paragraph',
        inlines: [{ text: 'link', link: 'https://example.com' }, { text: ' and ' }, { text: 'code', code: true }]
      }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const runs = doc.paragraphs[0].lines[0];
    expect(runs[0].link).toBe('https://example.com');
    expect(runs[2].fontName).toBeTruthy();
  });

  it('applies the Quote style to quote blocks', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [{ type: 'quote', inlines: [{ text: 'quoted' }] }], DEFAULT_INSERT_OPTIONS);
    expect(doc.paragraphs[0].styleBuiltIn).toBe('Quote');
  });

  it('inserts one Normal + monospace paragraph per code line', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [{ type: 'code', text: 'line one\nline two' }], DEFAULT_INSERT_OPTIONS);
    expect(doc.texts()).toEqual(['line one', 'line two']);
    expect(doc.paragraphs[0].lines[0][0].fontName).toBeTruthy();
    expect(doc.paragraphs[1].styleBuiltIn).toBe('Normal');
  });

  it('inserts a real empty Normal paragraph for a spacer block', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('One'), para(''), para('Two')], DEFAULT_INSERT_OPTIONS);
    expect(doc.texts()).toEqual(['One', '', 'Two']);
    expect(doc.paragraphs[1].styleBuiltIn).toBe('Normal');
  });

  it('applies a custom body style by name when configured', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [para('Hi')], { ...DEFAULT_INSERT_OPTIONS, bodyStyle: { kind: 'custom', name: 'NormalWeb' } });
    const [p] = doc.paragraphs;
    expect(p.styleName).toBe('NormalWeb');
    expect(p.styleBuiltIn).toBeUndefined();
    expect(report.mismatches).toEqual([]); // a custom style's styleBuiltIn ("Other") isn't compared
  });

  it('syncs once per plain block', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('Hi'), para('There')], DEFAULT_INSERT_OPTIONS);
    expect(doc.commitCount).toBe(2);
  });

  it('does nothing (and does not sync) for an empty block list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [], DEFAULT_INSERT_OPTIONS);
    expect(doc.commitCount).toBe(0);
    expect(report.steps).toEqual([]);
  });

  it('reports a clean read-back for a simple paste', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [para('One'), bullet('item'), para('Two')], DEFAULT_INSERT_OPTIONS);
    expect(report.mismatches).toEqual([]);
    expect(report.verifyError).toBeNull();
    expect(report.readBack?.paragraphs.map((p) => p.text)).toEqual(['One', 'item', 'Two']);
    expect(report.steps.map(formatStep)).toContain('block -1 verify: read-back matches');
  });
});

describe('insertBlocks — tables', () => {
  const table = (header: boolean): Block => ({
    type: 'table',
    header: header ? [[{ text: 'A' }], [{ text: 'B' }]] : null,
    rows: [[[{ text: '1' }], [{ text: '2' }]]]
  });

  it('inserts a table with plain-text values and the configured style', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [table(true)], DEFAULT_INSERT_OPTIONS);
    const t = doc.model.blocks.find(isTable);
    if (!t) throw new Error('expected table');
    expect(t.values).toEqual([
      ['A', 'B'],
      ['1', '2']
    ]);
    expect(t.style).toMatchObject({ styleBuiltIn: 'GridTable4', headerRowCount: 1, styleFirstColumn: true, styleBandedRows: true });
  });

  it('sets headerRowCount to 0 when the table has no header row', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [table(false)], DEFAULT_INSERT_OPTIONS);
    expect(doc.model.blocks.find(isTable)?.style?.headerRowCount).toBe(0);
  });

  it('leaves the empty cursor paragraph in place before a table that is the very first block', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [table(false)], DEFAULT_INSERT_OPTIONS);
    expect(doc.model.blocks.map((b) => b.kind)).toEqual(['paragraph', 'table']);
  });

  it('does not insert a stray empty paragraph before a table that follows other content', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('Intro'), table(false)], DEFAULT_INSERT_OPTIONS);
    expect(doc.model.blocks.map((b) => b.kind)).toEqual(['paragraph', 'table']);
  });

  it('does not add an extra empty paragraph after a table (only the house-style blank line)', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks = normalize([para('Intro'), table(true), para('After')], DEFAULT_NORMALIZE_OPTIONS);
    const report = await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const shape = doc.model.blocks.map((b) => (isParagraph(b) ? paragraphText(b) : '[table]'));
    expect(shape).toEqual(['Intro', '[table]', '', 'After']);
    expect(report.mismatches).toEqual([]);
  });

  it('puts a paragraph directly after a table when no blank line is wanted', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('Intro'), table(false), para('After')], DEFAULT_INSERT_OPTIONS);
    const shape = doc.model.blocks.map((b) => (isParagraph(b) ? paragraphText(b) : '[table]'));
    expect(shape).toEqual(['Intro', '[table]', 'After']);
  });

  it('separates two consecutive tables with one empty paragraph', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('Intro'), table(false), table(false)], DEFAULT_INSERT_OPTIONS);
    const shape = doc.model.blocks.map((b) => (isParagraph(b) ? paragraphText(b) : '[table]'));
    expect(shape).toEqual(['Intro', '[table]', '', '[table]']);
  });

  it('a list right after a table is still a proper list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [para('Intro'), table(false), bullet('x'), bullet('y')], DEFAULT_INSERT_OPTIONS);
    const [, x, y] = doc.paragraphs;
    expect(x.listRef?.list).toBeDefined();
    expect(y.listRef?.list).toBe(x.listRef?.list);
    expect(report.mismatches).toEqual([]);
  });
});

describe('insertBlocks — two-phase lists', () => {
  it('first inserts every block as a plain paragraph, then turns list runs into lists', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [para('Intro'), bullet('a'), bullet('b'), para('After')], DEFAULT_INSERT_OPTIONS);
    const firstListCall = doc.calls.indexOf('startNewList');
    const lastInsert = doc.calls.lastIndexOf('insertParagraphAfter');
    expect(firstListCall).toBeGreaterThan(lastInsert);
    expect(doc.calls).toContain('attachToList');
  });

  it('uses ListParagraph for list items, one shared list, and the right level for nested items', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [bullet('top'), bullet('nested', 1, 1), bullet('top again')], DEFAULT_INSERT_OPTIONS);
    const [top, nested, again] = doc.paragraphs;
    expect(top.styleBuiltIn).toBe('ListParagraph');
    expect(top.listRef?.level).toBe(0);
    expect(nested.listRef?.level).toBe(1);
    expect(again.listRef?.level).toBe(0);
    expect(nested.listRef?.list).toBe(top.listRef?.list);
    expect(again.listRef?.list).toBe(top.listRef?.list);
  });

  it("configures bullets and Word's bullet-button indents for every level used", async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [bullet('top'), bullet('nested', 1, 1)], DEFAULT_INSERT_OPTIONS);
    const list = doc.paragraphs[0].listRef!.list;
    expect([...list.bulletLevels].sort()).toEqual([0, 1]);
    expect(list.levelIndents.get(0)).toEqual({ textIndent: 36, bulletIndent: 18 });
    expect(list.levelIndents.get(1)).toEqual({ textIndent: 72, bulletIndent: 54 });
    const rb = report.readBack!.paragraphs;
    expect(rb[0]).toMatchObject({ leftIndent: 36, firstLineIndent: -18 });
    expect(rb[1]).toMatchObject({ leftIndent: 72, firstLineIndent: -18 });
  });

  it('gives two separate listIndex values two independent lists, each numbered from 1', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [numbered('a1'), numbered('a2'), para('between'), numbered('b1', 2), numbered('b2', 2)];
    const report = await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [a1, a2, , b1, b2] = doc.paragraphs;
    expect(a1.listRef!.list).not.toBe(b1.listRef!.list);
    expect([a1, a2, b1, b2].map(listString)).toEqual(['1.', '2.', '1.', '2.']);
    expect(report.runs.map((r) => r.mode)).toEqual(['list', 'list']);
    expect(report.mismatches).toEqual([]);
  });

  it('two adjacent separate lists with nothing between them stay separate lists', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [numbered('a1'), numbered('b1', 2), numbered('b2', 2)], DEFAULT_INSERT_OPTIONS);
    const [a1, b1, b2] = doc.paragraphs;
    expect(a1.listRef?.list).not.toBe(b1.listRef?.list);
    expect([a1, b1, b2].map(listString)).toEqual(['1.', '1.', '2.']);
    expect(report.mismatches).toEqual([]);
  });

  it('respects an explicit start number for the first item of an ordered list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [numbered('a', 1, 0, 5), numbered('b')], DEFAULT_INSERT_OPTIONS);
    const [a, b] = doc.paragraphs;
    expect(a.listRef?.list.startingNumbers.get(0)).toBe(5);
    expect([a, b].map(listString)).toEqual(['5.', '6.']);
  });

  it('never sets a starting number on a normal list that starts at 1', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [numbered('a'), numbered('b')], DEFAULT_INSERT_OPTIONS);
    expect(doc.calls).not.toContain('setLevelStartingNumber');
  });

  it('keeps going when a level-formatting call fails, and logs it', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('setLevelIndents', 99);
    const report = await insertBlocks(doc, [bullet('a'), bullet('b')], DEFAULT_INSERT_OPTIONS);
    const [a, b] = doc.paragraphs;
    expect(b.listRef?.list).toBe(a.listRef?.list);
    expect(report.steps.map(formatStep).some((s) => s.includes('setLevelIndents(0, 36, 18) failed'))).toBe(true);
  });
});

describe('insertBlocks — leaving a list cleanly', () => {
  it('a body paragraph right after a bulleted list is not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, [bullet('item'), para('After the list')], DEFAULT_INSERT_OPTIONS);
    const [, after] = doc.paragraphs;
    expect(after.listRef).toBeUndefined();
    expect(after.styleBuiltIn).toBe('Normal');
    expect(paragraphText(after)).toBe('After the list');
    expect(report.repaired).toBe(0);
  });

  it('a heading right after a list is not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [numbered('item'), { type: 'heading', level: 2, inlines: [{ text: 'Next section' }] }], DEFAULT_INSERT_OPTIONS);
    const [, heading] = doc.paragraphs;
    expect(heading.listRef).toBeUndefined();
    expect(heading.styleBuiltIn).toBe('Heading2');
  });

  it('a blank spacer paragraph right after a list is not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [bullet('item'), para(''), para('Body')], DEFAULT_INSERT_OPTIONS);
    const [, blank] = doc.paragraphs;
    expect(blank.listRef).toBeUndefined();
    expect(paragraphText(blank)).toBe('');
    expect(blank.styleBuiltIn).toBe('Normal');
  });

  it('detaches and restyles a reused cursor paragraph that was itself an empty list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '', anchorInList: true });
    const report = await insertBlocks(doc, [para('Not a bullet')], DEFAULT_INSERT_OPTIONS);
    const [p] = doc.paragraphs;
    expect(p.listRef).toBeUndefined();
    expect(p.styleBuiltIn).toBe('Normal');
    expect(report.mismatches).toEqual([]);
  });

  it('a fresh list on a reused cursor paragraph that was already a list item starts a new list at 1', async () => {
    const doc = new FakeWordDocument({ anchorText: '', anchorInList: true });
    const oldList = doc.anchor.listRef!.list;
    await insertBlocks(doc, [numbered('fresh item'), numbered('second')], DEFAULT_INSERT_OPTIONS);
    const [p, q] = doc.paragraphs;
    expect(p.listRef?.list).not.toBe(oldList);
    expect([p, q].map(listString)).toEqual(['1.', '2.']);
  });

  it('content after a non-empty list-item cursor does not inherit that list', async () => {
    const doc = new FakeWordDocument({ anchorText: 'existing bullet', anchorInList: true });
    const oldList = doc.anchor.listRef!.list;
    const report = await insertBlocks(doc, [para('Plain one'), para('Plain two'), bullet('new item')], DEFAULT_INSERT_OPTIONS);
    const [existing, one, two, item] = doc.paragraphs;
    expect(existing.listRef?.list).toBe(oldList);
    expect(one.listRef).toBeUndefined();
    expect(two.listRef).toBeUndefined();
    expect(item.listRef?.list).toBeDefined();
    expect(item.listRef?.list).not.toBe(oldList);
    expect(report.steps.map(formatStep).some((s) => s.includes('inherited from the list-item cursor'))).toBe(true);
    expect(report.mismatches).toEqual([]);
  });

  it('the repair pass detaches a paragraph that ended up a list item anyway', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const { paragraph } = await doc.getCursor();
    paragraph.startNewList();
    let styled = false;
    const repaired = await doc.verifyAndRepairListItems([
      {
        handle: paragraph,
        intendedListItem: false,
        reapplyStyle: () => {
          paragraph.setStyleBuiltIn('Normal');
          styled = true;
        }
      }
    ]);
    expect(repaired).toBe(1);
    expect(styled).toBe(true);
    expect(doc.paragraphs[0].listRef).toBeUndefined();
  });

  it('the repair pass leaves correct paragraphs alone', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const { paragraph } = await doc.getCursor();
    expect(await doc.verifyAndRepairListItems([{ handle: paragraph, intendedListItem: false }])).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Regression: what real Word for Mac 16.113 did with v1.0.1
// ---------------------------------------------------------------------------

/** Synthetic text with the exact shape of Kevin's failing paste. */
const MAC_REGRESSION_MARKDOWN = `## Section title

### First subsection

Intro paragraph for the first subsection.

### Second subsection

The following applies to the items below, namely for:

- Alpha one
- Alpha two
- Alpha three
- Alpha four
- Alpha five
- Alpha six

### Third subsection

A body paragraph in the third subsection.

This paragraph ends with a list of:

- Beta one
- Beta two
- Beta three
- Beta four
- Beta five
- Beta six
`;

function macRegressionBlocks(): Block[] {
  return normalize(parseMarkdown(MAC_REGRESSION_MARKDOWN), { ...DEFAULT_NORMALIZE_OPTIONS, topHeadingLevel: 1, blankLinesBetweenBlocks: true });
}

const ALPHA = ['Alpha one', 'Alpha two', 'Alpha three', 'Alpha four', 'Alpha five', 'Alpha six'];
const BETA = ['Beta one', 'Beta two', 'Beta three', 'Beta four', 'Beta five', 'Beta six'];

/** Replays v1.0.1's call order (reduced to what decides paragraph order) against the fake. */
async function replayV101(doc: FakeWordDocument, blocks: Block[]): Promise<void> {
  let cursor: WordParagraphHandle = (await doc.getCursor()).paragraph;
  let reusable = true;
  let escapeList: FakeList | null = null; // v1.0.1's `currentList`
  let run: { list: FakeList; listIndex: number } | null = null;

  const next = (text: string): WordParagraphHandle => {
    if (reusable) {
      reusable = false;
      cursor.insertRun(text, { bold: false, italic: false });
      return cursor;
    }
    const p = escapeList ? escapeList.insertParagraph(text, 'After') : cursor.insertParagraphAfter(text);
    escapeList = null;
    return p;
  };

  for (const block of blocks) {
    const text = blockPlainText(block);
    if (block.type === 'listItem') {
      if (run && run.listIndex === block.listIndex) {
        cursor = run.list.insertParagraph(text, 'End');
        continue;
      }
      const first = next(text) as FakeParagraphHandle;
      run = { list: first.startNewList(), listIndex: block.listIndex };
      cursor = first;
      continue;
    }
    if (run) {
      escapeList = run.list;
      run = null;
    }
    cursor = next(text);
  }
}

describe('regression: Word for Mac list.insertParagraph quirk', () => {
  it('the fake reproduces exactly what Word for Mac did with v1.0.1', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await replayV101(doc, macRegressionBlocks());
    const shape = doc.paragraphs.map((p) => {
      const t = paragraphText(p);
      return p.listRef ? `${t}*` : t;
    });
    // Kevin's observation: H1, H2, body, blank, H2, body, A1 (bulleted), blank, H2, body,
    // blank, body, B1 (bulleted), B6, B5, B4, B3, B2, A6, A5, A4, A3, A2.
    expect(shape).toEqual([
      'Section title',
      'First subsection',
      'Intro paragraph for the first subsection.',
      '',
      'Second subsection',
      'The following applies to the items below, namely for:',
      'Alpha one*',
      '',
      'Third subsection',
      'A body paragraph in the third subsection.',
      '',
      'This paragraph ends with a list of:',
      'Beta one*',
      'Beta six',
      'Beta five',
      'Beta four',
      'Beta three',
      'Beta two',
      'Alpha six',
      'Alpha five',
      'Alpha four',
      'Alpha three',
      'Alpha two'
    ]);
  });

  it('v1.0.2 keeps the order, and makes all 6 items of each list items of one shared list', async () => {
    const doc = new FakeWordDocument({ anchorText: '', trailingText: 'Existing text after the cursor.' });
    const report = await insertBlocks(doc, macRegressionBlocks(), DEFAULT_INSERT_OPTIONS);

    expect(doc.texts()).toEqual([
      'Section title',
      'First subsection',
      'Intro paragraph for the first subsection.',
      '',
      'Second subsection',
      'The following applies to the items below, namely for:',
      ...ALPHA,
      '',
      'Third subsection',
      'A body paragraph in the third subsection.',
      '',
      'This paragraph ends with a list of:',
      ...BETA,
      'Existing text after the cursor.'
    ]);

    const paras = doc.paragraphs;
    const byText = (t: string): FakeParagraph => paras.find((p) => paragraphText(p) === t)!;
    const alpha = ALPHA.map(byText);
    const beta = BETA.map(byText);

    for (const items of [alpha, beta]) {
      expect(items.every((p) => p.listRef !== undefined)).toBe(true);
      expect(new Set(items.map((p) => p.listRef!.list)).size).toBe(1);
      expect(items.every((p) => p.listRef!.level === 0 && p.styleBuiltIn === 'ListParagraph')).toBe(true);
      expect(items.map(listString)).toEqual(Array(6).fill('•'));
      expect(items[0].listRef!.list.members).toEqual(items);
    }
    expect(alpha[0].listRef!.list).not.toBe(beta[0].listRef!.list);

    // Non-list paragraphs after each list.
    const afterAlpha = paras[paras.indexOf(alpha[5]) + 1];
    expect(paragraphText(afterAlpha)).toBe('');
    expect(afterAlpha.listRef).toBeUndefined();
    const afterBeta = paras[paras.indexOf(beta[5]) + 1];
    expect(paragraphText(afterBeta)).toBe('Existing text after the cursor.');
    expect(afterBeta.listRef).toBeUndefined();
    expect(paras.filter((p) => p.listRef).length).toBe(12);

    // Headings and blank lines keep their styles.
    expect(byText('Section title').styleBuiltIn).toBe('Heading1');
    expect(byText('Third subsection').styleBuiltIn).toBe('Heading2');
    expect(paras.filter((p) => paragraphText(p) === '').every((p) => p.styleBuiltIn === 'Normal')).toBe(true);

    expect(report.runs.map((r) => [r.mode, r.itemCount])).toEqual([
      ['list', 6],
      ['list', 6]
    ]);
    expect(report.mismatches).toEqual([]);
    expect(report.repaired).toBe(0);
  });

  it('the read-back comparison catches the v1.0.1 result', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks = macRegressionBlocks();
    const good = await insertBlocks(new FakeWordDocument({ anchorText: '' }), blocks, DEFAULT_INSERT_OPTIONS);
    await replayV101(doc, blocks);
    const { compareReadBack } = await import('../src/word/verify');
    const mismatches = compareReadBack(good.expected, doc.paragraphs.map(readBackOf), good.runs);
    expect(mismatches[0]).toMatch(/^text order: paragraph #8: expected "Alpha two", found ""/);
    expect(mismatches.some((m) => m.includes('should be an item of list 1'))).toBe(true);
  });
});

describe('insertBlocks — per-item-lists fallback', () => {
  const sixNumbered = (): Block[] => [1, 2, 3, 4, 5, 6].map((n) => numbered(`item ${n}`));

  it('falls back when attachToList throws: every item is still a list item, numbered 1-6', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('attachToList');
    const report = await insertBlocks(doc, [...sixNumbered(), para('after')], DEFAULT_INSERT_OPTIONS);

    const items = doc.paragraphs.slice(0, 6);
    expect(items.every((p) => p.listRef)).toBe(true);
    expect(new Set(items.map((p) => p.listRef!.list)).size).toBe(6);
    expect(items.map(listString)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.']);
    expect(items[3].listRef!.list.levelIndents.get(0)).toEqual({ textIndent: 36, bulletIndent: 18 });
    expect(doc.paragraphs[6].listRef).toBeUndefined();

    expect(report.runs[0]).toMatchObject({ mode: 'per-item', fallbackFrom: 1 });
    const steps = report.steps.map(formatStep);
    expect(steps.some((s) => s.includes('list 1: per-item-lists fallback'))).toBe(true);
    expect(steps.some((s) => s.includes('attachToList threw (GeneralException @ Word.attachToList)'))).toBe(true);
    expect(report.mismatches).toEqual([]);
  });

  it('falls back when attachToList silently leaves the paragraph outside the list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.attachToListIsNoop = true;
    const report = await insertBlocks(doc, [bullet('a'), bullet('b'), bullet('c')], DEFAULT_INSERT_OPTIONS);
    const items = doc.paragraphs;
    expect(items.every((p) => p.listRef)).toBe(true);
    expect(items.map(listString)).toEqual(['•', '•', '•']);
    expect(report.runs[0]).toMatchObject({ mode: 'per-item', fallbackFrom: 1 });
    expect(report.steps.map(formatStep).some((s) => s.includes('attachToList left isListItem false'))).toBe(true);
  });

  it('keeps the items that did join, and falls back only for the rest', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('attachToList', 1, 1); // item 2 joins, item 3 fails
    const report = await insertBlocks(doc, sixNumbered(), DEFAULT_INSERT_OPTIONS);
    const items = doc.paragraphs;
    expect(items[1].listRef!.list).toBe(items[0].listRef!.list);
    expect(items[2].listRef!.list).not.toBe(items[0].listRef!.list);
    expect(items.map(listString)).toEqual(['1.', '2.', '3.', '4.', '5.', '6.']);
    expect(report.runs[0]).toMatchObject({ mode: 'per-item', fallbackFrom: 2 });
    expect(report.mismatches).toEqual([]);
  });

  it('falls back for the whole run when startNewList throws on the first item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('startNewList');
    const report = await insertBlocks(doc, [numbered('a'), numbered('b')], DEFAULT_INSERT_OPTIONS);
    expect(doc.paragraphs.every((p) => p.listRef)).toBe(true);
    expect(doc.paragraphs.map(listString)).toEqual(['1.', '2.']);
    expect(report.runs[0]).toMatchObject({ mode: 'per-item', fallbackFrom: 0 });
  });

  it('keeps nested levels and their numbering in the fallback', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.attachToListIsNoop = true;
    await insertBlocks(doc, [numbered('one'), numbered('one.a', 1, 1), numbered('one.b', 1, 1), numbered('two')], DEFAULT_INSERT_OPTIONS);
    const items = doc.paragraphs;
    expect(items.map((p) => p.listRef?.level)).toEqual([0, 1, 1, 0]);
    expect(items.map(listString)).toEqual(['1.', '1.', '2.', '2.']);
    expect(items[1].listRef!.list.levelIndents.get(1)).toEqual({ textIndent: 72, bulletIndent: 54 });
  });

  it('does not throw when even the fallback fails, but reports the missing list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('attachToList');
    doc.failNext('startNewList', 1, 1); // the first item's list works, item 2's own list fails
    const report = await insertBlocks(doc, [bullet('a'), bullet('b'), bullet('c')], DEFAULT_INSERT_OPTIONS);
    expect(doc.paragraphs.map((p) => Boolean(p.listRef))).toEqual([true, false, true]);
    expect(report.mismatches.some((m) => m.includes('"b": should be an item of list 1, but is not a list item'))).toBe(true);
    expect(report.steps.map(formatStep).some((s) => s.includes('item 2: startNewList failed'))).toBe(true);
  });
});

describe('insertBlocks — diagnostics', () => {
  it('a phase-1 failure throws an InsertError naming the block, type and action', async () => {
    const doc = new FakeWordDocument({ anchorText: 'existing' });
    const originalCommit = doc.commit.bind(doc);
    let calls = 0;
    doc.commit = async () => {
      calls++;
      if (calls === 1) throw new Error('simulated commit failure');
      return originalCommit();
    };

    const caught = await insertBlocks(doc, [{ type: 'heading', level: 1, inlines: [{ text: 'Title' }] }], DEFAULT_INSERT_OPTIONS).catch(
      (err: unknown) => err
    );
    expect(caught).toBeInstanceOf(InsertError);
    const err = caught as InsertError;
    expect(err.failedStep).toMatchObject({ blockIndex: 0, blockType: 'heading', action: 'insertParagraphAfter' });
    expect(err.steps.length).toBeGreaterThan(1);
  });

  it('reports the failing block when creating a later paragraph throws', async () => {
    const doc = new FakeWordDocument({ anchorText: 'existing' });
    doc.failNext('insertParagraphAfter', 1, 1);
    const caught = await insertBlocks(doc, [para('one'), bullet('two')], DEFAULT_INSERT_OPTIONS).catch((err: unknown) => err);
    expect(caught).toBeInstanceOf(InsertError);
    const err = caught as InsertError;
    expect(err.failedStep?.blockIndex).toBe(1);
    expect(err.failedStep?.blockType).toBe('listItem');
    expect(err.officeError).toMatchObject({ code: 'GeneralException', errorLocation: 'Word.insertParagraphAfter' });
  });

  it('a failing read-back is reported, not thrown', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('readBack');
    const report = await insertBlocks(doc, [para('one')], DEFAULT_INSERT_OPTIONS);
    expect(report.readBack).toBeNull();
    expect(report.verifyError).toContain('GeneralException');
    expect(report.steps.map(formatStep).some((s) => s.includes('readBack failed'))).toBe(true);
  });
});

describe('computeOrdinals', () => {
  it('counts per level, restarts deeper levels, and honours start numbers', () => {
    const items = [
      numbered('a', 1, 0, 3),
      numbered('a.1', 1, 1),
      numbered('a.2', 1, 1),
      numbered('b'),
      numbered('b.1', 1, 1),
      numbered('b.1.x', 1, 2, 7),
      numbered('c')
    ];
    expect(computeOrdinals(items)).toEqual([3, 1, 2, 4, 1, 7, 5]);
  });
});
