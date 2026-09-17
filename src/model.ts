/**
 * Intermediate block model.
 *
 * Both the Markdown parser and the HTML parser produce this shape, so
 * everything downstream (normalization, Word insertion, the preview panel)
 * only ever has to deal with one representation of "what got pasted".
 */

/** A run of text with (at most) simple, semantic-only formatting. */
export interface Inline {
  text: string;
  bold?: boolean;
  italic?: boolean;
  code?: boolean;
  link?: string;
}

export interface HeadingBlock {
  type: 'heading';
  /** 1-based heading level, pre-clamp. May exceed 9 before normalization. */
  level: number;
  inlines: Inline[];
}

export interface ParagraphBlock {
  type: 'paragraph';
  inlines: Inline[];
}

export interface ListItemBlock {
  type: 'listItem';
  ordered: boolean;
  /** 0-based nesting depth. */
  level: number;
  /**
   * Identifies which *separate* list this item belongs to. Two adjacent
   * ordered lists (e.g. separated by a paragraph) get different indices so
   * each one can be restarted at 1 in Word.
   */
  listIndex: number;
  /** Explicit start number for the first item of an ordered list, if any. */
  start?: number;
  inlines: Inline[];
}

export interface TableBlock {
  type: 'table';
  header: Inline[][] | null;
  rows: Inline[][][];
}

export interface QuoteBlock {
  type: 'quote';
  inlines: Inline[];
}

export interface CodeBlock {
  type: 'code';
  text: string;
}

export type Block =
  | HeadingBlock
  | ParagraphBlock
  | ListItemBlock
  | TableBlock
  | QuoteBlock
  | CodeBlock;

export function plainText(inlines: Inline[]): string {
  return inlines.map((i) => i.text).join('');
}

export function blockPlainText(block: Block): string {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
    case 'listItem':
    case 'quote':
      return plainText(block.inlines);
    case 'code':
      return block.text;
    case 'table': {
      const cells: string[] = [];
      if (block.header) cells.push(...block.header.map(plainText));
      for (const row of block.rows) cells.push(...row.map(plainText));
      return cells.join(' ');
    }
  }
}
