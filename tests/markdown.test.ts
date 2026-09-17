import { describe, it, expect } from 'vitest';
import { parseMarkdown } from '../src/parse/markdown';
import { plainText } from '../src/model';
import type { HeadingBlock, ListItemBlock, TableBlock, CodeBlock, QuoteBlock } from '../src/model';

describe('parseMarkdown', () => {
  it('parses headings at every level', () => {
    const blocks = parseMarkdown('# H1\n\n## H2\n\n### H3\n');
    expect(blocks.map((b) => b.type)).toEqual(['heading', 'heading', 'heading']);
    const [h1, h2, h3] = blocks as HeadingBlock[];
    expect(h1.level).toBe(1);
    expect(h2.level).toBe(2);
    expect(h3.level).toBe(3);
    expect(plainText(h1.inlines)).toBe('H1');
  });

  it('parses bold, italic, code and link runs', () => {
    const [block] = parseMarkdown('Some **bold** and *italic* and `code` and [link](https://example.com).');
    expect(block.type).toBe('paragraph');
    if (block.type !== 'paragraph') throw new Error('expected paragraph');
    const bold = block.inlines.find((i) => i.text === 'bold');
    const italic = block.inlines.find((i) => i.text === 'italic');
    const code = block.inlines.find((i) => i.text === 'code');
    const link = block.inlines.find((i) => i.text === 'link');
    expect(bold?.bold).toBe(true);
    expect(italic?.italic).toBe(true);
    expect(code?.code).toBe(true);
    expect(link?.link).toBe('https://example.com');
  });

  it('handles nested bullet lists with increasing level', () => {
    const blocks = parseMarkdown('- item 1\n  - nested item\n- item 2\n') as ListItemBlock[];
    expect(blocks).toHaveLength(3);
    expect(blocks[0]).toMatchObject({ level: 0, ordered: false });
    expect(blocks[1]).toMatchObject({ level: 1, ordered: false });
    expect(blocks[2]).toMatchObject({ level: 0, ordered: false });
    // nested item shares the same listIndex as its parent list
    expect(blocks[1].listIndex).toBe(blocks[0].listIndex);
    expect(blocks[2].listIndex).toBe(blocks[0].listIndex);
  });

  it('gives two separate ordered lists different listIndex values, each restarting at 1', () => {
    const md = '1. first\n2. second\n\nSome paragraph in between.\n\n1. new first\n2. new second\n';
    const blocks = parseMarkdown(md);
    const listItems = blocks.filter((b): b is ListItemBlock => b.type === 'listItem');
    expect(listItems).toHaveLength(4);
    const [a1, a2, b1, b2] = listItems;
    expect(a1.listIndex).toBe(a2.listIndex);
    expect(b1.listIndex).toBe(b2.listIndex);
    expect(a1.listIndex).not.toBe(b1.listIndex);
  });

  it('records an explicit start number only on the first item of an ordered list', () => {
    const blocks = parseMarkdown('5. fifth\n6. sixth\n') as ListItemBlock[];
    // marked treats this as one contiguous list (no interruption) so both share a listIndex;
    // CommonMark ignores subsequent numbers, only the first item's start is meaningful.
    expect(blocks[0].start).toBe(5);
    expect(blocks[1].start).toBeUndefined();
  });

  it('parses a GFM table with a header row', () => {
    const md = '| a | b |\n| --- | --- |\n| 1 | 2 |\n| 3 | 4 |\n';
    const [table] = parseMarkdown(md) as [TableBlock];
    expect(table.type).toBe('table');
    expect(table.header).not.toBeNull();
    expect(plainText(table.header![0])).toBe('a');
    expect(table.rows).toHaveLength(2);
    expect(plainText(table.rows[1][1])).toBe('4');
  });

  it('parses a blockquote', () => {
    const [q] = parseMarkdown('> a quote\n') as [QuoteBlock];
    expect(q.type).toBe('quote');
    expect(plainText(q.inlines)).toBe('a quote');
  });

  it('parses a fenced code block verbatim', () => {
    const [code] = parseMarkdown('```js\nconst x = 1;\nconsole.log(x);\n```\n') as [CodeBlock];
    expect(code.type).toBe('code');
    expect(code.text).toBe('const x = 1;\nconsole.log(x);');
  });

  it('drops horizontal rules', () => {
    const blocks = parseMarkdown('Before\n\n---\n\nAfter\n');
    expect(blocks.map((b) => b.type)).toEqual(['paragraph', 'paragraph']);
  });

  it('turns a hard line break into a "\\n" inline', () => {
    const [block] = parseMarkdown('Line one  \nLine two\n');
    if (block.type !== 'paragraph') throw new Error('expected paragraph');
    expect(block.inlines.some((i) => i.text === '\n')).toBe(true);
  });

  it('unescapes backslash escapes without treating them as markdown', () => {
    const [block] = parseMarkdown('Escaped \\*not bold\\* and \\# not heading.\n');
    if (block.type !== 'paragraph') throw new Error('expected paragraph');
    expect(plainText(block.inlines)).toContain('*not bold*');
    expect(block.inlines.every((i) => !i.bold)).toBe(true);
  });

  it('decodes HTML entities inside markdown text', () => {
    const [block] = parseMarkdown('Q&amp;A and 5 &lt; 6\n');
    if (block.type !== 'paragraph') throw new Error('expected paragraph');
    expect(plainText(block.inlines)).toBe('Q&A and 5 < 6');
  });

  it('handles a Dutch numbered heading and a bold label paragraph', () => {
    const blocks = parseMarkdown('## 1. Inleiding\n\n**Doelstelling:** dit project heeft als doel...\n');
    expect(blocks[0].type).toBe('heading');
    expect(plainText((blocks[0] as HeadingBlock).inlines)).toBe('1. Inleiding');
    if (blocks[1].type !== 'paragraph') throw new Error('expected paragraph');
    expect(blocks[1].inlines[0].bold).toBe(true);
    expect(blocks[1].inlines[0].text).toBe('Doelstelling:');
  });

  it('supports <p> content split across a list item plus a nested sub-list', () => {
    const md = '- item 1\n  - nested item\n- item 2\n';
    const items = parseMarkdown(md) as ListItemBlock[];
    expect(plainText(items[0].inlines)).toBe('item 1');
    expect(plainText(items[1].inlines)).toBe('nested item');
  });
});
