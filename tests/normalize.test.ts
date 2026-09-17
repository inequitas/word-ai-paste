import { describe, it, expect } from 'vitest';
import { normalize, DEFAULT_NORMALIZE_OPTIONS } from '../src/transform/normalize';
import { plainText } from '../src/model';
import type { Block, HeadingBlock, ParagraphBlock } from '../src/model';

function heading(level: number, text: string): HeadingBlock {
  return { type: 'heading', level, inlines: [{ text }] };
}
function para(text: string, opts: Partial<{ bold: boolean }> = {}): ParagraphBlock {
  return { type: 'paragraph', inlines: text ? [{ text, bold: opts.bold }] : [] };
}

describe('normalize — top heading level', () => {
  it('shifts headings so the highest input level maps to the configured top level, default 1', () => {
    const blocks: Block[] = [heading(2, 'A'), heading(3, 'B')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, boldLineToHeading: false, blankLinesBetweenBlocks: false });
    expect((result[0] as HeadingBlock).level).toBe(1);
    expect((result[1] as HeadingBlock).level).toBe(2);
  });

  it('shifts headings to a configured top level of 2, preserving relative depth', () => {
    const blocks: Block[] = [heading(1, 'A'), heading(2, 'B')];
    const result = normalize(blocks, {
      ...DEFAULT_NORMALIZE_OPTIONS,
      topHeadingLevel: 2,
      boldLineToHeading: false,
      blankLinesBetweenBlocks: false
    });
    expect((result[0] as HeadingBlock).level).toBe(2);
    expect((result[1] as HeadingBlock).level).toBe(3);
  });

  it('clamps shifted heading levels to 9', () => {
    const blocks: Block[] = [heading(1, 'A'), heading(9, 'B')];
    const result = normalize(blocks, {
      ...DEFAULT_NORMALIZE_OPTIONS,
      topHeadingLevel: 4,
      boldLineToHeading: false,
      blankLinesBetweenBlocks: false
    });
    expect((result[1] as HeadingBlock).level).toBe(9);
  });
});

describe('normalize — bold-only lines to headings', () => {
  it('promotes a standalone all-bold paragraph to a heading one level below the previous real heading', () => {
    const blocks: Block[] = [heading(1, 'Chapter'), para('Pseudo heading', { bold: true }), para('Body text')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(result[1].type).toBe('heading');
    expect((result[1] as HeadingBlock).level).toBe(2);
  });

  it('uses the configured top heading level when there is no preceding real heading', () => {
    const blocks: Block[] = [para('Pseudo heading', { bold: true })];
    const result = normalize(blocks, {
      ...DEFAULT_NORMALIZE_OPTIONS,
      topHeadingLevel: 3,
      blankLinesBetweenBlocks: false
    });
    expect((result[0] as HeadingBlock).level).toBe(3);
  });

  it('strips a single trailing colon from a promoted heading', () => {
    const blocks: Block[] = [para('Doelstelling:', { bold: true })];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as HeadingBlock).inlines)).toBe('Doelstelling');
  });

  it('does not promote a paragraph that mixes bold and non-bold text', () => {
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: 'Bold ', bold: true }, { text: 'and plain' }] }];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(result[0].type).toBe('paragraph');
  });

  it('does not promote a long bold paragraph (> ~100 chars)', () => {
    const long = 'x'.repeat(120);
    const blocks: Block[] = [para(long, { bold: true })];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(result[0].type).toBe('paragraph');
  });

  it('can be disabled', () => {
    const blocks: Block[] = [para('Pseudo heading', { bold: true })];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, boldLineToHeading: false, blankLinesBetweenBlocks: false });
    expect(result[0].type).toBe('paragraph');
  });
});

describe('normalize — citation markers', () => {
  it('removes 【…】 markers', () => {
    const blocks: Block[] = [para('Some fact【1†source】continues.')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).toBe('Some fact continues.');
  });

  it('removes private-use-area citation tokens', () => {
    const puaToken = String.fromCharCode(0xe200, 0xe201, 0xe202);
    const blocks: Block[] = [para(`Fact${puaToken}here.`)];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).toBe('Fact here.');
  });

  it('does not strip normal bracketed references like [1]', () => {
    const blocks: Block[] = [para('See reference [1] for details.')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).toBe('See reference [1] for details.');
  });

  it('can be disabled', () => {
    const blocks: Block[] = [para('Some fact【1†source】continues.')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, removeCitations: false, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).toContain('【1†source】');
  });
});

describe('normalize — manual heading numbers', () => {
  it('is off by default', () => {
    const blocks: Block[] = [heading(1, '2.3 Doelstelling')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as HeadingBlock).inlines)).toBe('2.3 Doelstelling');
  });

  it('strips "1.", "2.3" and "2.3.1" style prefixes when enabled', () => {
    const blocks: Block[] = [heading(1, '1. Inleiding'), heading(2, '2.3 Doelstelling'), heading(3, '2.3.1 Sub')];
    const result = normalize(blocks, {
      ...DEFAULT_NORMALIZE_OPTIONS,
      removeManualHeadingNumbers: true,
      blankLinesBetweenBlocks: false
    });
    expect(plainText((result[0] as HeadingBlock).inlines)).toBe('Inleiding');
    expect(plainText((result[1] as HeadingBlock).inlines)).toBe('Doelstelling');
    expect(plainText((result[2] as HeadingBlock).inlines)).toBe('Sub');
  });
});

describe('normalize — emoji', () => {
  it('is off by default', () => {
    const blocks: Block[] = [para('Great work! 🎉')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).toContain('🎉');
  });

  it('strips emoji when enabled', () => {
    const blocks: Block[] = [para('Great work! 🎉')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, removeEmoji: true, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).not.toContain('🎉');
  });
});

describe('normalize — always-on cleanup', () => {
  it('collapses internal whitespace and trims block edges', () => {
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: '  Hello    world  ' }] }];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(plainText((result[0] as ParagraphBlock).inlines)).toBe('Hello world');
  });

  it('drops empty paragraphs', () => {
    const blocks: Block[] = [para('Real text'), { type: 'paragraph', inlines: [] }, para('More text')];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(result).toHaveLength(2);
  });

  it('drops leading and trailing blank blocks', () => {
    const blocks: Block[] = [{ type: 'paragraph', inlines: [{ text: '   ' }] }, para('Middle'), { type: 'paragraph', inlines: [] }];
    const result = normalize(blocks, { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
    expect(result).toHaveLength(1);
    expect(plainText((result[0] as ParagraphBlock).inlines)).toBe('Middle');
  });
});

describe('normalize — blank lines between blocks (house style)', () => {
  const opts = { ...DEFAULT_NORMALIZE_OPTIONS, boldLineToHeading: false, blankLinesBetweenBlocks: true };

  it('inserts a blank paragraph between two body paragraphs', () => {
    const blocks: Block[] = [para('One'), para('Two')];
    const result = normalize(blocks, opts);
    expect(result.map((b) => b.type)).toEqual(['paragraph', 'paragraph', 'paragraph']);
    expect((result[1] as ParagraphBlock).inlines).toHaveLength(0);
  });

  it('inserts nothing between a body paragraph and the list that follows it', () => {
    const blocks: Block[] = [para('Intro:'), { type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'item' }] }];
    const result = normalize(blocks, opts);
    expect(result).toHaveLength(2);
  });

  it('inserts a blank paragraph between a list and the body paragraph that follows it', () => {
    const blocks: Block[] = [{ type: 'listItem', ordered: false, level: 0, listIndex: 1, inlines: [{ text: 'item' }] }, para('After')];
    const result = normalize(blocks, opts);
    expect(result).toHaveLength(3);
    expect(result[1].type).toBe('paragraph');
    expect((result[1] as ParagraphBlock).inlines).toHaveLength(0);
  });

  it('inserts a blank paragraph before any heading', () => {
    const blocks: Block[] = [para('Body'), heading(2, 'Next section')];
    const result = normalize(blocks, opts);
    expect(result).toHaveLength(3);
    expect(result[1].type).toBe('paragraph');
    expect((result[1] as ParagraphBlock).inlines).toHaveLength(0);
  });

  it('inserts nothing after a heading', () => {
    const blocks: Block[] = [heading(1, 'Title'), para('Body')];
    const result = normalize(blocks, opts);
    expect(result).toHaveLength(2);
  });

  it('inserts a blank paragraph after a table when more content follows', () => {
    const blocks: Block[] = [{ type: 'table', header: null, rows: [[[{ text: 'a' }]]] }, para('After table')];
    const result = normalize(blocks, opts);
    expect(result).toHaveLength(3);
    expect(result[1].type).toBe('paragraph');
    expect((result[1] as ParagraphBlock).inlines).toHaveLength(0);
  });

  it('inserts nothing when the option is off', () => {
    const blocks: Block[] = [para('One'), para('Two')];
    const result = normalize(blocks, { ...opts, blankLinesBetweenBlocks: false });
    expect(result).toHaveLength(2);
  });
});
