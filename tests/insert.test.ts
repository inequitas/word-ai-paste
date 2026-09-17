import { describe, it, expect } from 'vitest';
import { insertBlocks, DEFAULT_INSERT_OPTIONS, InsertError } from '../src/word/insert';
import type { Block } from '../src/model';
import { FakeWordDocument, isParagraph, isTable, paragraphText } from './fakeWordAdapter';

describe('insertBlocks', () => {
  it('reuses an empty cursor paragraph for the first block instead of creating a new one', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: 'Hello' }] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);

    expect(doc.model.blocks).toHaveLength(1);
    const p = doc.model.blocks[0];
    if (!isParagraph(p)) throw new Error('expected paragraph');
    expect(paragraphText(p)).toBe('Hello');
    expect(p.styleBuiltIn).toBe('Normal');
  });

  it('inserts after a non-empty cursor paragraph instead of overwriting it', async () => {
    const doc = new FakeWordDocument({ anchorText: 'existing text' });
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: 'New' }] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);

    expect(doc.model.blocks).toHaveLength(2);
    const [first, second] = doc.model.blocks;
    if (!isParagraph(first) || !isParagraph(second)) throw new Error('expected paragraphs');
    expect(paragraphText(first)).toBe('existing text');
    expect(paragraphText(second)).toBe('New');
  });

  it('sets styleBuiltIn to HeadingN, clamped at 9', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'heading', level: 2, inlines: [{ text: 'A' }] },
      { type: 'heading', level: 15, inlines: [{ text: 'B' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [a, b] = doc.model.blocks;
    if (!isParagraph(a) || !isParagraph(b)) throw new Error('expected paragraphs');
    expect(a.styleBuiltIn).toBe('Heading2');
    expect(b.styleBuiltIn).toBe('Heading9');
  });

  it('always sets bold/italic explicitly (true and false) on every run', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'paragraph', inlines: [{ text: 'bold ', bold: true }, { text: 'plain' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [p] = doc.model.blocks;
    if (!isParagraph(p)) throw new Error('expected paragraph');
    const runs = p.lines[0];
    expect(runs[0]).toMatchObject({ text: 'bold ', bold: true, italic: false });
    expect(runs[1]).toMatchObject({ text: 'plain', bold: false, italic: false });
  });

  it('sets a hyperlink on a linked run and a monospace font on a code run', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      {
        type: 'paragraph',
        inlines: [
          { text: 'link', link: 'https://example.com' },
          { text: ' and ' },
          { text: 'code', code: true }
        ]
      }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [p] = doc.model.blocks;
    if (!isParagraph(p)) throw new Error('expected paragraph');
    const runs = p.lines[0];
    expect(runs[0].link).toBe('https://example.com');
    expect(runs[2].fontName).toBeTruthy();
  });

  it('applies the Quote style to quote blocks', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [{ type: 'quote', inlines: [{ text: 'quoted' }] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [p] = doc.model.blocks;
    if (!isParagraph(p)) throw new Error('expected paragraph');
    expect(p.styleBuiltIn).toBe('Quote');
  });

  it('inserts one Normal + monospace paragraph per code line', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [{ type: 'code', text: 'line one\nline two' }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    expect(doc.model.blocks).toHaveLength(2);
    const [l1, l2] = doc.model.blocks;
    if (!isParagraph(l1) || !isParagraph(l2)) throw new Error('expected paragraphs');
    expect(paragraphText(l1)).toBe('line one');
    expect(paragraphText(l2)).toBe('line two');
    expect(l1.lines[0][0].fontName).toBeTruthy();
  });

  it('uses the ListParagraph style for list items and sets the level on nested items', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'top' }] },
      { type: 'listItem', ordered: false, level: 1, listIndex: 1, inlines: [{ text: 'nested' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [top, nested] = doc.model.blocks;
    if (!isParagraph(top) || !isParagraph(nested)) throw new Error('expected paragraphs');
    expect(top.styleBuiltIn).toBe('ListParagraph');
    expect(top.listRef?.level).toBe(0);
    expect(nested.listRef?.level).toBe(1);
    // both items belong to the same underlying Word list
    expect(nested.listRef?.list).toBe(top.listRef?.list);
  });

  it('gives two separate listIndex values two independent Word lists, each restarting at 1', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'a1' }], start: undefined },
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'a2' }] },
      { type: 'paragraph', inlines: [{ text: 'between' }] },
      { type: 'listItem', ordered: true, level: 0, listIndex: 2, inlines: [{ text: 'b1' }] },
      { type: 'listItem', ordered: true, level: 0, listIndex: 2, inlines: [{ text: 'b2' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    const listA = paras[0].listRef!.list;
    const listB = paras[3].listRef!.list;
    expect(listA).not.toBe(listB);
    expect(listA.numberLevels.get(0)).toBe(1);
    expect(listB.numberLevels.get(0)).toBe(1);
  });

  it('respects an explicit start number for the first item of an ordered list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'a' }], start: 5 }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [p] = doc.model.blocks.filter(isParagraph);
    expect(p.listRef?.list.numberLevels.get(0)).toBe(5);
  });

  it('inserts a table with plain-text values and the configured style, no direct header bold', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      {
        type: 'table',
        header: [[{ text: 'A' }], [{ text: 'B' }]],
        rows: [[[{ text: '1' }], [{ text: '2' }]]]
      }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    // The pre-existing empty cursor paragraph is left in place before the
    // table (see the dedicated "leading blank paragraph" test below) — the
    // table itself is whichever block is a table.
    const t = doc.model.blocks.find(isTable);
    if (!t) throw new Error('expected table');
    expect(t.values).toEqual([
      ['A', 'B'],
      ['1', '2']
    ]);
    expect(t.style?.styleBuiltIn).toBe('GridTable4');
    expect(t.style?.headerRowCount).toBe(1);
    expect(t.style?.styleFirstColumn).toBe(true);
    expect(t.style?.styleBandedRows).toBe(true);
  });

  it('sets headerRowCount to 0 when the table has no header row, regardless of the option default', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [{ type: 'table', header: null, rows: [[[{ text: '1' }], [{ text: '2' }]]] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const t = doc.model.blocks.find(isTable);
    if (!t) throw new Error('expected table');
    expect(t.style?.headerRowCount).toBe(0);
  });

  it('leaves the empty cursor paragraph in place before a table that is the very first inserted block', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [{ type: 'table', header: null, rows: [[[{ text: '1' }]]] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    expect(doc.model.blocks).toHaveLength(2);
    expect(doc.model.blocks[0].kind).toBe('paragraph');
    expect(doc.model.blocks[1].kind).toBe('table');
  });

  it('does not insert a stray empty paragraph before a table that follows other content', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'paragraph', inlines: [{ text: 'Intro' }] },
      { type: 'table', header: null, rows: [[[{ text: '1' }]]] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    expect(doc.model.blocks).toHaveLength(2);
    expect(doc.model.blocks[0].kind).toBe('paragraph');
    expect(doc.model.blocks[1].kind).toBe('table');
  });

  it('inserts a real empty paragraph for a spacer block (empty inlines) without runs', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'paragraph', inlines: [{ text: 'One' }] },
      { type: 'paragraph', inlines: [] },
      { type: 'paragraph', inlines: [{ text: 'Two' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras).toHaveLength(3);
    expect(paragraphText(paras[1])).toBe('');
    expect(paras[1].styleBuiltIn).toBe('Normal');
  });

  it('applies a custom body style by name when configured', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: 'Hi' }] }];
    await insertBlocks(doc, blocks, { ...DEFAULT_INSERT_OPTIONS, bodyStyle: { kind: 'custom', name: 'NormalWeb' } });
    const [p] = doc.model.blocks;
    if (!isParagraph(p)) throw new Error('expected paragraph');
    expect(p.styleName).toBe('NormalWeb');
    expect(p.styleBuiltIn).toBeUndefined();
  });

  it('commits exactly once', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [{ type: 'paragraph', inlines: [{ text: 'Hi' }] }], DEFAULT_INSERT_OPTIONS);
    expect(doc.commitCount).toBe(1);
  });

  it('does nothing (and does not commit) for an empty block list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    await insertBlocks(doc, [], DEFAULT_INSERT_OPTIONS);
    expect(doc.commitCount).toBe(0);
  });
});

describe('insertBlocks — leaving a list cleanly', () => {
  // The fake adapter reproduces Word's real quirk (insertParagraphAfter on
  // a list item continues the list), so these tests fail against a naive
  // implementation and pass only because insert.ts routes non-list content
  // after a list through WordListHandle.insertParagraphAfter() instead.

  it('a body paragraph right after a bulleted list is not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'item' }] },
      { type: 'paragraph', inlines: [{ text: 'After the list' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras).toHaveLength(2);
    expect(paras[1].listRef).toBeUndefined();
    expect(paras[1].styleBuiltIn).toBe('Normal');
    expect(paragraphText(paras[1])).toBe('After the list');
    expect(doc.repairedCount).toBe(0); // the primary fix should handle this, not the safety net
  });

  it('a heading right after a list is not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'item' }] },
      { type: 'heading', level: 2, inlines: [{ text: 'Next section' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras[1].listRef).toBeUndefined();
    expect(paras[1].styleBuiltIn).toBe('Heading2');
  });

  it('a blank spacer paragraph right after a list is not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'item' }] },
      { type: 'paragraph', inlines: [] },
      { type: 'paragraph', inlines: [{ text: 'Body' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras).toHaveLength(3);
    expect(paras[1].listRef).toBeUndefined();
    expect(paragraphText(paras[1])).toBe('');
    expect(paras[1].styleBuiltIn).toBe('Normal');
  });

  it('nested list items still end up outside the list once a non-list block follows', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'top' }] },
      { type: 'listItem', ordered: false, level: 1, listIndex: 1, inlines: [{ text: 'nested' }] },
      { type: 'paragraph', inlines: [{ text: 'After' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras[2].listRef).toBeUndefined();
    expect(paragraphText(paras[2])).toBe('After');
  });

  it('two adjacent separate lists with nothing between them both stay correctly list-owned', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'a1' }] },
      { type: 'listItem', ordered: true, level: 0, listIndex: 2, inlines: [{ text: 'b1' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras[0].listRef?.list).not.toBe(paras[1].listRef?.list);
    expect(paras[1].listRef?.list.numberLevels.get(0)).toBe(1);
  });

  it('detaches and restyles a reused cursor paragraph that was itself an empty list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '', anchorInList: true });
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: 'Not a bullet' }] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [p] = doc.model.blocks.filter(isParagraph);
    expect(p.listRef).toBeUndefined();
    expect(p.styleBuiltIn).toBe('Normal');
    expect(paragraphText(p)).toBe('Not a bullet');
  });

  it('starting a fresh list on a reused cursor paragraph that was already a (different) list item restarts at 1', async () => {
    const doc = new FakeWordDocument({ anchorText: '', anchorInList: true });
    const blocks: Block[] = [{ type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'fresh item' }] }];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const [p] = doc.model.blocks.filter(isParagraph);
    expect(p.listRef?.level).toBe(0);
    expect(p.listRef?.list.numberLevels.get(0)).toBe(1);
  });

  it('a table right after a list resets list tracking for whatever follows it', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'item' }] },
      { type: 'table', header: null, rows: [[[{ text: '1' }]]] },
      { type: 'paragraph', inlines: [{ text: 'After table' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    const last = paras[paras.length - 1];
    expect(last.listRef).toBeUndefined();
    expect(paragraphText(last)).toBe('After table');
  });
});

describe('WordDocument.verifyAndRepairListItems (fake adapter safety net)', () => {
  it('detaches and restyles a paragraph that ended up a list item when it should not have', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const { paragraph } = await doc.getCursor();
    const list = paragraph.startNewList();
    list.ensureBulletLevel(0);
    let styled = false;
    await doc.verifyAndRepairListItems([
      {
        handle: paragraph,
        intendedListItem: false,
        reapplyStyle: () => {
          paragraph.setStyleBuiltIn('Normal');
          styled = true;
        }
      }
    ]);
    expect(doc.repairedCount).toBe(1);
    expect(styled).toBe(true);
    const [p] = doc.model.blocks.filter(isParagraph);
    expect(p.listRef).toBeUndefined();
    expect(p.styleBuiltIn).toBe('Normal');
  });

  it('logs a warning (without touching anything) when an intended list item is not one', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const { paragraph } = await doc.getCursor();
    await doc.verifyAndRepairListItems([{ handle: paragraph, intendedListItem: true }]);
    expect(doc.repairedCount).toBe(0);
    expect(doc.warnings).toHaveLength(1);
  });

  it('does nothing for correctly-matched records', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const { paragraph } = await doc.getCursor();
    await doc.verifyAndRepairListItems([{ handle: paragraph, intendedListItem: false }]);
    expect(doc.repairedCount).toBe(0);
    expect(doc.warnings).toHaveLength(0);
  });
});

describe('insertBlocks — OOXML fallback when the list API throws', () => {
  it('falls back to insertOoxmlList (replacing the anchor) when startNewList() throws on the first item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('startNewList');
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'a' }] },
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'b' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);

    expect(doc.ooxmlCalls).toHaveLength(1);
    expect(doc.ooxmlCalls[0].anchorIsEmpty).toBe(true);
    expect(doc.ooxmlCalls[0].items).toHaveLength(2); // the whole run, since nothing had committed yet

    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras).toHaveLength(2);
    expect(paras.every((p) => p.viaOoxml)).toBe(true);
    expect(paras.every((p) => p.styleBuiltIn === 'ListParagraph')).toBe(true);
    expect(paragraphText(paras[0])).toBe('a');
    expect(paragraphText(paras[1])).toBe('b');
  });

  it('falls back only for the remaining items when a later item fails to join the list', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('insertItemAfter'); // fails on item 2, after item 1 already succeeded
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'first' }] },
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'second' }] },
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'third' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);

    expect(doc.ooxmlCalls).toHaveLength(1);
    expect(doc.ooxmlCalls[0].anchorIsEmpty).toBe(false); // item 1 already exists; insert after it
    expect(doc.ooxmlCalls[0].items).toHaveLength(2); // only items 2 and 3

    const paras = doc.model.blocks.filter(isParagraph);
    expect(paras).toHaveLength(3);
    expect(paragraphText(paras[0])).toBe('first');
    expect(paras[0].viaOoxml).toBeUndefined(); // created via the primary API path
    expect(paragraphText(paras[1])).toBe('second');
    expect(paras[1].viaOoxml).toBe(true);
    expect(paragraphText(paras[2])).toBe('third');
    expect(paras[2].viaOoxml).toBe(true);
  });

  it('a non-list block right after an ooxml-fallback list is still not a list item', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('startNewList');
    const blocks: Block[] = [
      { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'a' }] },
      { type: 'paragraph', inlines: [{ text: 'after' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    const paras = doc.model.blocks.filter(isParagraph);
    const last = paras[paras.length - 1];
    expect(paragraphText(last)).toBe('after');
    expect(last.listRef).toBeUndefined();
  });

  it('records both the failure and the fallback in the step log without throwing', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('startNewList');
    const blocks: Block[] = [{ type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'a' }] }];
    await expect(insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS)).resolves.toBeUndefined();
  });

  it('still restarts numbering at 1 for an ordered list rebuilt via the ooxml fallback', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('startNewList');
    const blocks: Block[] = [
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'a' }] },
      { type: 'listItem', ordered: true, level: 0, listIndex: 1, inlines: [{ text: 'b' }] }
    ];
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    expect(doc.ooxmlCalls[0].items.every((i) => i.ordered)).toBe(true);
  });
});

describe('insertBlocks — diagnostics on total failure', () => {
  it('throws an InsertError carrying the step log and the failing step when both paths fail', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('startNewList');
    const originalInsertOoxmlList = doc.insertOoxmlList.bind(doc);
    doc.insertOoxmlList = async () => {
      throw new Error('ooxml also failed');
    };
    void originalInsertOoxmlList;

    const blocks: Block[] = [{ type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'a' }] }];

    let caught: unknown;
    try {
      await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    } catch (err) {
      caught = err;
    }

    expect(caught).toBeInstanceOf(InsertError);
    const err = caught as InsertError;
    expect(err.steps.length).toBeGreaterThan(0);
    expect(err.steps.some((s) => s.action.includes('ooxml-fallback'))).toBe(true);
    expect(err.message).toContain('block 0');
  });

  it('a plain non-list failure reports the block index/type/action that failed', async () => {
    const doc = new FakeWordDocument({ anchorText: 'existing' });
    const originalCommit = doc.commit.bind(doc);
    let calls = 0;
    doc.commit = async () => {
      calls++;
      if (calls === 1) throw new Error('simulated commit failure');
      return originalCommit();
    };
    const blocks: Block[] = [{ type: 'heading', level: 1, inlines: [{ text: 'Title' }] }];

    let caught: unknown;
    try {
      await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);
    } catch (err) {
      caught = err;
    }

    expect(caught).toBeInstanceOf(InsertError);
    const err = caught as InsertError;
    expect(err.failedStep).toMatchObject({ blockIndex: 0, blockType: 'heading', action: 'insertParagraphAfter' });
  });
});
