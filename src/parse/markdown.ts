import { lexer } from 'marked';
import type { Token, Tokens, TokensList } from 'marked';
import type { Block, Inline } from '../model';
import { decodeEntities } from './entities';

/**
 * Parse a Markdown string into our block model using marked's lexer
 * (tokens only — we never touch marked's HTML renderer, so nothing here
 * ever produces or executes HTML).
 */
export function parseMarkdown(src: string): Block[] {
  const tokens: TokensList = lexer(src, { gfm: true, breaks: false });
  const blocks: Block[] = [];
  let listCounter = 0;

  for (const tok of tokens) {
    convertBlockToken(tok, blocks, () => ++listCounter);
  }

  return blocks;
}

function convertBlockToken(tok: Token, blocks: Block[], nextListIndex: () => number): void {
  switch (tok.type) {
    case 'heading': {
      const t = tok as Tokens.Heading;
      blocks.push({ type: 'heading', level: t.depth, inlines: inlineTokens(t.tokens) });
      break;
    }
    case 'paragraph': {
      const t = tok as Tokens.Paragraph;
      const inlines = inlineTokens(t.tokens);
      if (inlines.length) blocks.push({ type: 'paragraph', inlines });
      break;
    }
    case 'blockquote': {
      const t = tok as Tokens.Blockquote;
      const inlines = flattenToInlines(t.tokens);
      blocks.push({ type: 'quote', inlines });
      break;
    }
    case 'code': {
      const t = tok as Tokens.Code;
      blocks.push({ type: 'code', text: t.text });
      break;
    }
    case 'table': {
      const t = tok as Tokens.Table;
      const header = t.header.length ? t.header.map((cell) => inlineTokens(cell.tokens)) : null;
      const rows = t.rows.map((row) => row.map((cell) => inlineTokens(cell.tokens)));
      blocks.push({ type: 'table', header, rows });
      break;
    }
    case 'list': {
      const idx = nextListIndex();
      emitList(tok as Tokens.List, 0, idx, blocks);
      break;
    }
    case 'space':
    case 'hr':
    case 'html':
    case 'def':
      // hr: explicitly dropped per spec. space/def/raw html blocks carry no
      // renderable content we support.
      break;
    default: {
      // Unknown block type: best-effort fallback so we never silently lose
      // an entire section of the paste.
      const text = (tok as { text?: string }).text;
      if (text) blocks.push({ type: 'paragraph', inlines: [{ text: decodeEntities(text) }] });
    }
  }
}

function emitList(listTok: Tokens.List, level: number, listIndex: number, blocks: Block[]): void {
  let isFirst = true;
  for (const item of listTok.items) {
    const { inlines, subLists } = splitListItemTokens(item.tokens);
    blocks.push({
      type: 'listItem',
      ordered: listTok.ordered,
      level,
      listIndex,
      start: listTok.ordered && isFirst && typeof listTok.start === 'number' && listTok.start > 0 ? listTok.start : undefined,
      inlines
    });
    isFirst = false;
    for (const sub of subLists) {
      emitList(sub, level + 1, listIndex, blocks);
    }
  }
}

/** Splits a list item's child tokens into its own inline content plus any nested sub-lists. */
function splitListItemTokens(tokens: Token[]): { inlines: Inline[]; subLists: Tokens.List[] } {
  const inlines: Inline[] = [];
  const subLists: Tokens.List[] = [];

  for (const t of tokens) {
    if (t.type === 'list') {
      subLists.push(t as Tokens.List);
      continue;
    }
    if (t.type === 'text' || t.type === 'paragraph') {
      const withTokens = t as Tokens.Text | Tokens.Paragraph;
      const parts = withTokens.tokens ? inlineTokens(withTokens.tokens) : [{ text: decodeEntities(withTokens.text) }];
      appendWithBreak(inlines, parts);
      continue;
    }
    // Rare block content inside a list item (nested blockquote/code/table):
    // flatten to plain text rather than dropping it.
    const text = (t as { text?: string }).text;
    if (text) appendWithBreak(inlines, [{ text: decodeEntities(text) }]);
  }

  return { inlines, subLists };
}

function appendWithBreak(target: Inline[], parts: Inline[]): void {
  if (!parts.length) return;
  if (target.length) target.push({ text: '\n' });
  target.push(...parts);
}

/** Flattens a blockquote's (possibly multi-paragraph) token tree into one inline run. */
function flattenToInlines(tokens: Token[]): Inline[] {
  const out: Inline[] = [];
  for (const t of tokens) {
    if (t.type === 'paragraph') {
      appendWithBreak(out, inlineTokens((t as Tokens.Paragraph).tokens));
    } else if (t.type === 'blockquote') {
      appendWithBreak(out, flattenToInlines((t as Tokens.Blockquote).tokens));
    } else if (t.type === 'list') {
      // Simplify a list nested inside a quote to plain lines; full list
      // fidelity inside quotes is out of scope for v1.
      const listTok = t as Tokens.List;
      for (const item of listTok.items) {
        const { inlines } = splitListItemTokens(item.tokens);
        appendWithBreak(out, inlines);
      }
    } else {
      const text = (t as { text?: string }).text;
      if (text) appendWithBreak(out, [{ text: decodeEntities(text) }]);
    }
  }
  return out;
}

type InlineCtx = Pick<Inline, 'bold' | 'italic' | 'code' | 'link'>;

function inlineTokens(tokens: Token[] | undefined, ctx: InlineCtx = {}): Inline[] {
  if (!tokens) return [];
  const out: Inline[] = [];

  for (const t of tokens) {
    switch (t.type) {
      case 'text': {
        const tt = t as Tokens.Text;
        if (tt.tokens && tt.tokens.length) {
          out.push(...inlineTokens(tt.tokens, ctx));
        } else {
          out.push(makeInline(decodeEntities(tt.text), ctx));
        }
        break;
      }
      case 'strong':
        out.push(...inlineTokens((t as Tokens.Strong).tokens, { ...ctx, bold: true }));
        break;
      case 'em':
        out.push(...inlineTokens((t as Tokens.Em).tokens, { ...ctx, italic: true }));
        break;
      case 'del':
        // Strikethrough isn't part of the model; keep the text, drop the styling.
        out.push(...inlineTokens((t as Tokens.Del).tokens, ctx));
        break;
      case 'codespan':
        out.push(makeInline((t as Tokens.Codespan).text, { ...ctx, code: true }));
        break;
      case 'link':
        out.push(...inlineTokens((t as Tokens.Link).tokens, { ...ctx, link: (t as Tokens.Link).href }));
        break;
      case 'image': {
        const img = t as Tokens.Image;
        const alt = img.text || img.title || '';
        if (alt) out.push(makeInline(alt, ctx));
        break;
      }
      case 'br':
        out.push(makeInline('\n', ctx));
        break;
      case 'escape':
        out.push(makeInline((t as Tokens.Escape).text, ctx));
        break;
      case 'html': {
        const raw = (t as Tokens.HTML).raw.trim();
        if (/^<br\s*\/?>$/i.test(raw)) out.push(makeInline('\n', ctx));
        break;
      }
      default: {
        const text = (t as { text?: string }).text;
        if (text) out.push(makeInline(decodeEntities(text), ctx));
      }
    }
  }

  return out;
}

function makeInline(text: string, ctx: InlineCtx): Inline {
  const inline: Inline = { text };
  if (ctx.bold) inline.bold = true;
  if (ctx.italic) inline.italic = true;
  if (ctx.code) inline.code = true;
  if (ctx.link) inline.link = ctx.link;
  return inline;
}
