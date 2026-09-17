import { describe, it, expect } from 'vitest';
import { insertBlocks, DEFAULT_INSERT_OPTIONS } from '../src/word/insert';
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
