import type { Block, Inline } from '../model';

export interface NormalizeOptions {
  /** Shift headings so the input's highest heading level maps to this (1-9, UI offers 1-4). */
  topHeadingLevel: number;
  /** Promote a standalone all-bold paragraph line to a heading. */
  boldLineToHeading: boolean;
  /** Strip 【…】 citation markers and ChatGPT's private-use-area citation tokens. */
  removeCitations: boolean;
  /** Strip manual heading numbers like "1.", "2.3", "2.3.1" from the start of headings. */
  removeManualHeadingNumbers: boolean;
  /** Strip emoji from all text. */
  removeEmoji: boolean;
  /**
   * Kevin's house style: a blank Normal paragraph between most block
   * transitions instead of relying on paragraph spacing. See
   * `insertBlankLines` for the exact rules.
   */
  blankLinesBetweenBlocks: boolean;
}

export const DEFAULT_NORMALIZE_OPTIONS: NormalizeOptions = {
  topHeadingLevel: 1,
  boldLineToHeading: true,
  removeCitations: true,
  removeManualHeadingNumbers: false,
  removeEmoji: false,
  blankLinesBetweenBlocks: true
};

const MAX_HEADING_LEVEL = 9;
const BOLD_HEADING_MAX_CHARS = 100;

export function normalize(blocks: Block[], options: NormalizeOptions = DEFAULT_NORMALIZE_OPTIONS): Block[] {
  let result = shiftHeadingLevels(blocks, options.topHeadingLevel);

  if (options.boldLineToHeading) {
    result = promoteBoldLinesToHeadings(result, options.topHeadingLevel);
  }
  if (options.removeCitations) {
    result = result.map((b) => mapBlockText(b, removeCitationMarkers));
  }
  if (options.removeManualHeadingNumbers) {
    result = stripManualHeadingNumbers(result);
  }
  if (options.removeEmoji) {
    result = result.map((b) => mapBlockText(b, stripEmoji));
  }

  // Always: trim/collapse whitespace, drop empty blocks (also covers
  // "drop horizontal rules" and "strip leading/trailing blank blocks" —
  // both parsers never emit an hr block at all, and every genuinely blank
  // block is dropped here regardless of where it sits).
  result = cleanupBlocks(result);

  if (options.blankLinesBetweenBlocks) {
    result = insertBlankLines(result);
  }

  return result;
}

// ---------------------------------------------------------------------------
// 1. Top heading level
// ---------------------------------------------------------------------------

function shiftHeadingLevels(blocks: Block[], topHeadingLevel: number): Block[] {
  const levels = blocks.filter((b): b is Extract<Block, { type: 'heading' }> => b.type === 'heading').map((b) => b.level);
  if (!levels.length) return blocks;
  const shift = clamp(topHeadingLevel, 1, MAX_HEADING_LEVEL) - Math.min(...levels);
  if (shift === 0) return blocks;
  return blocks.map((b) => (b.type === 'heading' ? { ...b, level: clamp(b.level + shift, 1, MAX_HEADING_LEVEL) } : b));
}

function clamp(n: number, min: number, max: number): number {
  return Math.min(Math.max(n, min), max);
}

// ---------------------------------------------------------------------------
// 2. Bold-only lines -> headings
// ---------------------------------------------------------------------------

function promoteBoldLinesToHeadings(blocks: Block[], topHeadingLevel: number): Block[] {
  let lastRealHeadingLevel: number | null = null;

  return blocks.map((block) => {
    if (block.type === 'heading') {
      lastRealHeadingLevel = block.level;
      return block;
    }
    if (block.type !== 'paragraph') return block;

    const nonEmpty = block.inlines.filter((i) => i.text.trim().length > 0);
    if (!nonEmpty.length) return block;
    if (nonEmpty.some((i) => i.text === '\n')) return block; // multi-line: not a single heading-ish line
    if (!nonEmpty.every((i) => i.bold)) return block;

    const plain = nonEmpty.map((i) => i.text).join('').trim();
    if (!plain || plain.length > BOLD_HEADING_MAX_CHARS) return block;

    const level = lastRealHeadingLevel != null ? Math.min(lastRealHeadingLevel + 1, MAX_HEADING_LEVEL) : clamp(topHeadingLevel, 1, MAX_HEADING_LEVEL);

    const inlines = nonEmpty.map((i) => ({ ...i }));
    const lastIndex = inlines.length - 1;
    inlines[lastIndex] = { ...inlines[lastIndex], text: inlines[lastIndex].text.replace(/:\s*$/, '') };

    return { type: 'heading', level, inlines };
  });
}

// ---------------------------------------------------------------------------
// 3. Citation markers
// ---------------------------------------------------------------------------

function removeCitationMarkers(text: string): string {
  // Replaced with a space rather than deleted outright: these markers often
  // sit right up against the surrounding words with no space of their own
  // (e.g. "fact【1†source】more"), and the later whitespace-collapse pass
  // cleans up any resulting double spaces.
  return text
    .replace(/【[^】]*】/g, ' ')
    // ChatGPT encodes citation tokens (e.g. "cite turn0search0") as a run of
    // private-use-area characters that render as invisible/tofu glyphs.
    .replace(/[-]+/g, ' ');
}

// ---------------------------------------------------------------------------
// 4. Manual heading numbers
// ---------------------------------------------------------------------------

function stripManualHeadingNumbers(blocks: Block[]): Block[] {
  return blocks.map((block) => {
    if (block.type !== 'heading' || !block.inlines.length) return block;
    const [first, ...rest] = block.inlines;
    const newText = first.text.replace(/^\s*\d+(?:\.\d+)*\.?\s+/, '');
    if (newText === first.text) return block;
    return { ...block, inlines: [{ ...first, text: newText }, ...rest] };
  });
}

// ---------------------------------------------------------------------------
// 5. Emoji
// ---------------------------------------------------------------------------

function stripEmoji(text: string): string {
  return text.replace(/\p{Extended_Pictographic}/gu, '').replace(/[‍️]/g, '');
}

// ---------------------------------------------------------------------------
// Always: whitespace + empty-block cleanup
// ---------------------------------------------------------------------------

function cleanupBlocks(blocks: Block[]): Block[] {
  return blocks.map(cleanBlock).filter((b) => !isBlockEmpty(b));
}

function cleanBlock(block: Block): Block {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
    case 'listItem':
    case 'quote':
      return { ...block, inlines: cleanInlines(block.inlines) };
    case 'table':
      return {
        ...block,
        header: block.header ? block.header.map(cleanInlines) : null,
        rows: block.rows.map((row) => row.map(cleanInlines))
      };
    case 'code':
      return { ...block, text: cleanCodeText(block.text) };
  }
}

function isBlockEmpty(block: Block): boolean {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
    case 'listItem':
    case 'quote':
      return block.inlines.length === 0;
    case 'code':
      return block.text.trim().length === 0;
    case 'table':
      return (!block.header || block.header.length === 0) && block.rows.length === 0;
  }
}

function cleanInlines(inlines: Inline[]): Inline[] {
  let cleaned = inlines.map((i) => ({ ...i, text: collapseWhitespace(i.text) })).filter((i) => i.text.length > 0);

  if (cleaned.length) {
    cleaned[0] = { ...cleaned[0], text: cleaned[0].text.replace(/^ +/, '') };
    const lastIdx = cleaned.length - 1;
    cleaned[lastIdx] = { ...cleaned[lastIdx], text: cleaned[lastIdx].text.replace(/ +$/, '') };
  }
  cleaned = cleaned.filter((i) => i.text.length > 0);

  while (cleaned.length && cleaned[0].text === '\n') cleaned.shift();
  while (cleaned.length && cleaned[cleaned.length - 1].text === '\n') cleaned.pop();

  return cleaned;
}

function collapseWhitespace(text: string): string {
  if (text === '\n') return text; // an intentional hard line break marker
  return text.replace(/\s+/g, ' ');
}

function cleanCodeText(text: string): string {
  const lines = text.split('\n');
  while (lines.length && lines[0].trim() === '') lines.shift();
  while (lines.length && lines[lines.length - 1].trim() === '') lines.pop();
  return lines.join('\n');
}

function mapBlockText(block: Block, fn: (s: string) => string): Block {
  switch (block.type) {
    case 'heading':
    case 'paragraph':
    case 'listItem':
    case 'quote':
      return { ...block, inlines: mapInlineText(block.inlines, fn) };
    case 'table':
      return {
        ...block,
        header: block.header ? block.header.map((cell) => mapInlineText(cell, fn)) : null,
        rows: block.rows.map((row) => row.map((cell) => mapInlineText(cell, fn)))
      };
    case 'code':
      return block; // never touch code content
  }
}

function mapInlineText(inlines: Inline[], fn: (s: string) => string): Inline[] {
  return inlines.map((i) => (i.text === '\n' ? i : { ...i, text: fn(i.text) }));
}

// ---------------------------------------------------------------------------
// 6. Blank lines between blocks (Kevin's house style)
// ---------------------------------------------------------------------------

type BlockKind = 'heading' | 'body' | 'list' | 'table';

function kindOf(block: Block): BlockKind {
  switch (block.type) {
    case 'heading':
      return 'heading';
    case 'listItem':
      return 'list';
    case 'table':
      return 'table';
    case 'paragraph':
    case 'quote':
    case 'code':
      return 'body';
  }
}

/**
 * Rules (Kevin's house style — spacing via empty paragraphs, not paragraph
 * spacing settings):
 *   body -> body: blank · list -> body: blank · body -> list: nothing
 *   anything -> heading: blank · heading -> anything: nothing
 *   after a table (if more content follows): blank
 * Checked in that priority order so every case above resolves unambiguously
 * (e.g. heading -> heading correctly yields "nothing", matching
 * "heading -> anything: nothing" taking precedence over "anything -> heading").
 */
function needsBlankBetween(prev: Block, next: Block): boolean {
  const p = kindOf(prev);
  const n = kindOf(next);
  if (p === 'heading') return false;
  if (n === 'heading') return true;
  if (p === 'table') return true;
  if (p === 'body' && n === 'body') return true;
  if (p === 'list' && n === 'body') return true;
  return false;
}

function insertBlankLines(blocks: Block[]): Block[] {
  if (blocks.length < 2) return blocks;
  const out: Block[] = [blocks[0]];
  for (let i = 1; i < blocks.length; i++) {
    const prev = blocks[i - 1];
    const next = blocks[i];
    if (needsBlankBetween(prev, next)) {
      out.push({ type: 'paragraph', inlines: [] });
    }
    out.push(next);
  }
  return out;
}
