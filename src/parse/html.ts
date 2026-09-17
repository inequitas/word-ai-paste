import type { Block, Inline, TableBlock } from '../model';
import { decodeEntities } from './entities';

/**
 * Parse pasted HTML (chat-tool "rendered" HTML, or Word's own HTML export)
 * into our block model.
 *
 * All class/style attributes are ignored except `font-weight` (bold) and
 * `font-style` (italic) on inline elements — everything else (colors,
 * fonts, sizes, margins…) is deliberately dropped so the document's own
 * styles govern the result. The one exception is macOS's "Cocoa HTML
 * Writer" output (what `textutil`/Cocoa's rich-text-to-HTML conversion —
 * and so the system clipboard's HTML flavor for a plain RTF chat reply —
 * produces): it never emits `<h1>`-`<h6>`, only `<p class="pN">` paragraphs
 * with sizes recorded in a `<style>` block. See `computeHeadingSizeRank`.
 */
export function parseHtml(html: string): Block[] {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  // Read the stylesheet before removeUnwanted() strips the <style> tag.
  const state = createState(doc);
  removeUnwanted(doc);
  const out: Block[] = [];
  walkChildren(doc.body, out, state);
  return out;
}

interface ParseState {
  nextListIndex: () => number;
  msoListIndexById: Map<string, number>;
  /** CSS class name -> font-size in px, parsed from any <style> block (Cocoa HTML Writer). */
  classFontSizePx: Map<string, number>;
  /** font-size in px -> heading level (1 = largest), for sizes bigger than the body text size. */
  headingSizeRank: Map<number, number>;
}

function createState(doc: Document): ParseState {
  let counter = 0;
  const classFontSizePx = parseStyleSheet(doc);
  const headingSizeRank = computeHeadingSizeRank(doc, classFontSizePx);
  return {
    nextListIndex: () => ++counter,
    msoListIndexById: new Map(),
    classFontSizePx,
    headingSizeRank
  };
}

/** Parses `<style>` blocks for simple `.class { font-size: ... }` / `font: SIZEpx family` rules. */
function parseStyleSheet(doc: Document): Map<string, number> {
  const map = new Map<string, number>();
  doc.querySelectorAll('style').forEach((styleEl) => {
    const css = styleEl.textContent || '';
    const ruleRe = /([^{}]+)\{([^{}]*)\}/g;
    let m: RegExpExecArray | null;
    while ((m = ruleRe.exec(css))) {
      const size = extractFontSizePx(m[2]);
      if (size == null) continue;
      for (const sel of m[1].split(',')) {
        const classMatch = /\.([A-Za-z0-9_-]+)/.exec(sel.trim());
        if (classMatch) map.set(classMatch[1], size);
      }
    }
  });
  return map;
}

function extractFontSizePx(declarations: string): number | null {
  const fsMatch = /font-size\s*:\s*([\d.]+)px/i.exec(declarations);
  if (fsMatch) return parseFloat(fsMatch[1]);
  // Cocoa's `font` shorthand, e.g. "font: 24.0px Times".
  const fontMatch = /(?:^|;)\s*font\s*:\s*([^;]+)/i.exec(declarations);
  if (fontMatch) {
    const sizeInShorthand = /([\d.]+)px/.exec(fontMatch[1]);
    if (sizeInShorthand) return parseFloat(sizeInShorthand[1]);
  }
  return null;
}

function elementFontSizePx(el: Element, classFontSizePx: Map<string, number>): number | null {
  const inlineSize = extractFontSizePx(el.getAttribute('style') || '');
  if (inlineSize != null) return inlineSize;
  for (const c of (el.getAttribute('class') || '').split(/\s+/)) {
    if (c && classFontSizePx.has(c)) return classFontSizePx.get(c)!;
  }
  return null;
}

/**
 * Body size = the font size covering the most characters of paragraph/list
 * text in the document (weighted by text length, not paragraph count).
 * Every distinct size bigger than that is ranked into a heading level,
 * largest first. Returns an empty map when the HTML has no such sizing
 * info at all (the ordinary case — real <h1>-<h6> / Mso classes handle
 * that HTML instead, see walkBlockElement).
 */
function computeHeadingSizeRank(doc: Document, classFontSizePx: Map<string, number>): Map<number, number> {
  const charsBySize = new Map<number, number>();
  doc.querySelectorAll('p, li').forEach((el) => {
    const size = elementFontSizePx(el, classFontSizePx);
    if (size == null) return;
    const text = (el.textContent || '').trim();
    if (!text) return;
    charsBySize.set(size, (charsBySize.get(size) || 0) + text.length);
  });
  if (!charsBySize.size) return new Map();

  let bodySize = 0;
  let bodyChars = -1;
  for (const [size, chars] of charsBySize) {
    if (chars > bodyChars) {
      bodyChars = chars;
      bodySize = size;
    }
  }

  const rank = new Map<number, number>();
  Array.from(charsBySize.keys())
    .filter((size) => size > bodySize)
    .sort((a, b) => b - a)
    .forEach((size, i) => rank.set(size, i + 1));
  return rank;
}

/** True when the element has some non-whitespace text and ALL of it is bold. */
function isAllBold(el: Element): boolean {
  let hasText = false;
  let allBold = true;

  const walk = (node: Node, bold: boolean): void => {
    if (!allBold) return;
    if (node.nodeType === Node.TEXT_NODE) {
      const text = (node.textContent || '').trim();
      if (text) {
        hasText = true;
        if (!bold) allBold = false;
      }
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const elx = node as Element;
    const tag = elx.tagName.toLowerCase();
    if (STRIP_TAGS.has(tag) || isHiddenElement(elx)) return;
    const nextBold = bold || tag === 'b' || tag === 'strong' || isStyleBold(elx);
    elx.childNodes.forEach((c) => walk(c, nextBold));
  };

  walk(el, false);
  return hasText && allBold;
}

function isStyleBold(el: Element): boolean {
  const style = el.getAttribute('style') || '';
  const fw = /font-weight\s*:\s*([^;]+)/i.exec(style)?.[1]?.trim().toLowerCase();
  return fw === 'bold' || fw === 'bolder' || (fw !== undefined && /^\d+$/.test(fw) && parseInt(fw, 10) >= 600);
}

const STRIP_TAGS = new Set(['script', 'style', 'svg', 'button', 'noscript', 'template']);

const BLOCK_TAGS = new Set([
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'p', 'div', 'ul', 'ol', 'li', 'table',
  'thead', 'tbody', 'tfoot', 'tr', 'td', 'th',
  'blockquote', 'pre', 'hr',
  'section', 'article', 'header', 'footer', 'figure', 'main'
]);

const BLOCK_CHILD_SELECTOR = 'p,div,h1,h2,h3,h4,h5,h6,ul,ol,table,blockquote,pre,hr';

function removeUnwanted(doc: Document): void {
  doc.querySelectorAll(Array.from(STRIP_TAGS).join(',')).forEach((n) => n.remove());
  // Elements that are visually hidden (screen-reader-only helper text,
  // display:none chrome, etc.) — drop the whole subtree up front.
  doc.querySelectorAll('*').forEach((el) => {
    if (isHiddenElement(el)) el.remove();
  });
}

function isHiddenElement(el: Element): boolean {
  if (el.hasAttribute('hidden')) return true;
  if ((el.getAttribute('aria-hidden') || '').toLowerCase() === 'true') return true;
  const classes = (el.getAttribute('class') || '').split(/\s+/).map((c) => c.toLowerCase());
  if (classes.some((c) => ['sr-only', 'visually-hidden', 'visuallyhidden', 'screen-reader-only', 'screen-reader-text'].includes(c))) {
    return true;
  }
  const style = (el.getAttribute('style') || '').toLowerCase();
  if (/display\s*:\s*none/.test(style)) return true;
  if (/visibility\s*:\s*hidden/.test(style)) return true;
  if (/position\s*:\s*absolute/.test(style) && /width\s*:\s*1px/.test(style) && /height\s*:\s*1px/.test(style)) return true;
  return false;
}

// ---------------------------------------------------------------------------
// Block-level walking
// ---------------------------------------------------------------------------

function walkChildren(parent: Node, out: Block[], state: ParseState): void {
  const children = Array.from(parent.childNodes);
  let buffer: Inline[] = [];

  const flush = () => {
    if (buffer.length) {
      out.push({ type: 'paragraph', inlines: buffer });
      buffer = [];
    }
  };

  for (let i = 0; i < children.length; i++) {
    const node = children[i];

    if (node.nodeType === Node.TEXT_NODE) {
      const txt = (node.textContent ?? '').replace(/\s+/g, ' ');
      if (txt.trim()) buffer.push(makeInline(txt, {}));
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) continue;

    const el = node as Element;
    if (isHiddenElement(el)) continue;
    const tag = el.tagName.toLowerCase();
    if (STRIP_TAGS.has(tag)) continue;

    if (!BLOCK_TAGS.has(tag)) {
      // Inline content sitting directly among block-level siblings.
      processInlineChild(el, {}, buffer);
      continue;
    }

    flush();

    // Heuristic: a short <div> label-or-toolbar element immediately before a
    // fenced code block (chat tools wrap each code block in a header bar
    // with a language tag and a "Copy code" button) — treat it as chat-UI
    // chrome and drop it rather than emitting a stray one-word paragraph.
    // Scoped to <div> only so it never swallows a genuine semantic block
    // (blockquote, heading, list, table, …) that merely happens to sit
    // right before a <pre>.
    if (tag === 'div' && !el.querySelector('pre')) {
      const next = children.slice(i + 1).find((n) => n.nodeType === Node.ELEMENT_NODE) as Element | undefined;
      if (next && (next.tagName.toLowerCase() === 'pre' || next.querySelector?.('pre'))) {
        const text = (el.textContent ?? '').trim();
        if (text.length > 0 && text.length <= 60 && !text.includes('\n')) continue;
      }
    }

    walkBlockElement(el, out, state);
  }

  flush();
}

function walkBlockElement(el: Element, out: Block[], state: ParseState): void {
  const tag = el.tagName.toLowerCase();

  const msoHeadingLevel = getMsoHeadingLevel(el);
  if (/^h[1-6]$/.test(tag) || msoHeadingLevel) {
    const level = msoHeadingLevel ?? Number(tag[1]);
    const inlines = inlineFromChildren(el, {});
    if (inlines.length) out.push({ type: 'heading', level, inlines });
    return;
  }

  // Cocoa HTML Writer output (macOS's RTF-to-HTML, e.g. `textutil`, which is
  // what a "formatted text" chat reply looks like once it hits the system
  // clipboard) never uses <h1>-<h6>: a heading is just a bold paragraph in a
  // larger-than-body font size, recorded only via a <style> class. Real
  // heading tags / Mso classes above always take precedence over this.
  if (tag === 'p' && state.headingSizeRank.size > 0 && isAllBold(el)) {
    const size = elementFontSizePx(el, state.classFontSizePx);
    const level = size != null ? state.headingSizeRank.get(size) : undefined;
    if (level) {
      const inlines = inlineFromChildren(el, {});
      if (inlines.length) {
        out.push({ type: 'heading', level, inlines });
        return;
      }
    }
  }

  if (tag === 'p' || tag === 'div') {
    const listInfo = getMsoListInfo(el, state);
    if (listInfo) {
      const inlines = inlineFromChildren(el, {});
      out.push({
        type: 'listItem',
        ordered: listInfo.ordered,
        level: listInfo.level,
        listIndex: listInfo.listIndex,
        inlines
      });
      return;
    }
    if (tag === 'div' && hasBlockChild(el)) {
      walkChildren(el, out, state);
      return;
    }
    const inlines = inlineFromChildren(el, {});
    if (inlines.length) out.push({ type: 'paragraph', inlines });
    return;
  }

  if (tag === 'ul' || tag === 'ol') {
    const idx = state.nextListIndex();
    emitHtmlList(el, 0, idx, out);
    return;
  }

  if (tag === 'table') {
    out.push(tableFromElement(el));
    return;
  }

  if (tag === 'blockquote') {
    const inlines = flattenBlockquoteEl(el);
    if (inlines.length) out.push({ type: 'quote', inlines });
    return;
  }

  if (tag === 'pre') {
    out.push({ type: 'code', text: extractCodeText(el) });
    return;
  }

  if (tag === 'hr') return;

  // section/article/header/footer/figure/main/etc: transparent container.
  walkChildren(el, out, state);
}

function hasBlockChild(el: Element): boolean {
  return el.querySelector(BLOCK_CHILD_SELECTOR) !== null;
}

// ---------------------------------------------------------------------------
// Lists
// ---------------------------------------------------------------------------

function emitHtmlList(listEl: Element, level: number, listIndex: number, out: Block[]): void {
  const ordered = listEl.tagName.toLowerCase() === 'ol';
  const start = ordered ? parseStartAttr(listEl) : undefined;
  let isFirst = true;

  for (const child of Array.from(listEl.children)) {
    if (child.tagName.toLowerCase() !== 'li') continue;
    const { inlines, subLists } = collectListItemContent(child);
    out.push({
      type: 'listItem',
      ordered,
      level,
      listIndex,
      start: ordered && isFirst ? start : undefined,
      inlines
    });
    isFirst = false;
    for (const sub of subLists) {
      emitHtmlList(sub, level + 1, listIndex, out);
    }
  }
}

function parseStartAttr(el: Element): number | undefined {
  const n = parseInt(el.getAttribute('start') || '', 10);
  return Number.isFinite(n) && n > 0 ? n : undefined;
}

function collectListItemContent(li: Element): { inlines: Inline[]; subLists: Element[] } {
  const inlines: Inline[] = [];
  const subLists: Element[] = [];
  let line: Inline[] = [];

  const flush = () => {
    if (line.length) {
      appendWithBreak(inlines, line);
      line = [];
    }
  };

  for (const node of Array.from(li.childNodes)) {
    if (node.nodeType === Node.ELEMENT_NODE) {
      const el = node as Element;
      const tag = el.tagName.toLowerCase();
      if (isHiddenElement(el) || STRIP_TAGS.has(tag)) continue;
      if (tag === 'ul' || tag === 'ol') {
        flush();
        subLists.push(el);
        continue;
      }
      if (tag === 'p' || tag === 'div') {
        flush();
        appendWithBreak(inlines, inlineFromChildren(el, {}));
        continue;
      }
    }
    processInlineChild(node, {}, line);
  }
  flush();

  return { inlines, subLists };
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

function tableFromElement(tableEl: Element): TableBlock {
  const rows: { el: Element; fromThead: boolean }[] = [];
  for (const child of Array.from(tableEl.children)) {
    const tag = child.tagName.toLowerCase();
    if (tag === 'tr') {
      rows.push({ el: child, fromThead: false });
    } else if (tag === 'thead' || tag === 'tbody' || tag === 'tfoot') {
      for (const tr of Array.from(child.children)) {
        if (tr.tagName.toLowerCase() === 'tr') rows.push({ el: tr, fromThead: tag === 'thead' });
      }
    }
  }

  let header: Inline[][] | null = null;
  let start = 0;
  if (rows.length) {
    const firstRowCells = Array.from(rows[0].el.children).filter((c) => ['td', 'th'].includes(c.tagName.toLowerCase()));
    const hasThCell = firstRowCells.some((c) => c.tagName.toLowerCase() === 'th');
    // Cocoa HTML Writer tables have no <thead>/<th> at all — the header row
    // is just a first row whose cells are all bold (see parseHtml's doc
    // comment).
    const allCellsBold = firstRowCells.length > 0 && firstRowCells.every((c) => isAllBold(c));
    if (rows[0].fromThead || hasThCell || allCellsBold) {
      header = cellsOf(rows[0].el);
      start = 1;
    }
  }

  const bodyRows = rows.slice(start).map((r) => cellsOf(r.el));
  return { type: 'table', header, rows: bodyRows };
}

function cellsOf(tr: Element): Inline[][] {
  return Array.from(tr.children)
    .filter((c) => ['td', 'th'].includes(c.tagName.toLowerCase()))
    .map((c) => inlineFromChildren(c, {}));
}

// ---------------------------------------------------------------------------
// Blockquote / code
// ---------------------------------------------------------------------------

function flattenBlockquoteEl(el: Element): Inline[] {
  const out: Inline[] = [];
  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === Node.ELEMENT_NODE) {
      const c = child as Element;
      const tag = c.tagName.toLowerCase();
      if (isHiddenElement(c) || STRIP_TAGS.has(tag)) continue;
      if (tag === 'p' || tag === 'div') {
        appendWithBreak(out, inlineFromChildren(c, {}));
        continue;
      }
      if (tag === 'blockquote') {
        appendWithBreak(out, flattenBlockquoteEl(c));
        continue;
      }
      if (tag === 'ul' || tag === 'ol') {
        for (const li of Array.from(c.children)) {
          if (li.tagName.toLowerCase() !== 'li') continue;
          const { inlines } = collectListItemContent(li);
          appendWithBreak(out, inlines);
        }
        continue;
      }
    }
    const tmp: Inline[] = [];
    processInlineChild(child, {}, tmp);
    appendWithBreak(out, tmp);
  }
  return out;
}

function extractCodeText(preEl: Element): string {
  const codeEl = preEl.querySelector('code');
  const raw = (codeEl ?? preEl).textContent ?? '';
  return raw.replace(/^\n/, '').replace(/\n$/, '');
}

// ---------------------------------------------------------------------------
// Word ("Mso*") HTML export support
// ---------------------------------------------------------------------------

function getMsoHeadingLevel(el: Element): number | null {
  const cls = el.getAttribute('class') || '';
  const m = /\bMso[Hh]eading(\d)\b/.exec(cls) || /\bHeading(\d)\b/.exec(cls);
  if (m) {
    const n = parseInt(m[1], 10);
    if (n >= 1 && n <= 9) return n;
  }
  return null;
}

function getMsoListInfo(
  el: Element,
  state: ParseState
): { ordered: boolean; level: number; listIndex: number } | null {
  const cls = el.getAttribute('class') || '';
  const style = el.getAttribute('style') || '';
  const isListParagraph = /\bMsoListParagraph(CxSp(First|Middle|Last))?\b/i.test(cls);
  const m = /mso-list\s*:\s*l(\d+)\s+level(\d+)/i.exec(style);
  if (!isListParagraph && !m) return null;

  const listId = m ? `l${m[1]}` : 'l0';
  const level = m ? Math.max(0, parseInt(m[2], 10) - 1) : 0;

  let ordered = false;
  const markerSpan = Array.from(el.querySelectorAll('span')).find((s) =>
    /mso-list\s*:\s*ignore/i.test(s.getAttribute('style') || '')
  );
  if (markerSpan) {
    const markerText = (markerSpan.textContent || '').trim();
    ordered = /^[0-9]+[.)]/.test(markerText) || /^[a-zA-Z]+[.)]/.test(markerText);
    markerSpan.remove();
  }

  if (!state.msoListIndexById.has(listId)) {
    state.msoListIndexById.set(listId, state.nextListIndex());
  }
  const listIndex = state.msoListIndexById.get(listId)!;

  return { ordered, level, listIndex };
}

// ---------------------------------------------------------------------------
// Inline-level walking
// ---------------------------------------------------------------------------

type InlineCtx = Pick<Inline, 'bold' | 'italic' | 'code' | 'link'>;

function inlineFromChildren(node: Node, ctx: InlineCtx): Inline[] {
  const out: Inline[] = [];
  node.childNodes.forEach((child) => processInlineChild(child, ctx, out));
  return out;
}

function processInlineChild(child: Node, ctx: InlineCtx, out: Inline[]): void {
  if (child.nodeType === Node.TEXT_NODE) {
    const collapsed = (child.textContent ?? '').replace(/\s+/g, ' ');
    if (collapsed) out.push(makeInline(collapsed, ctx));
    return;
  }
  if (child.nodeType !== Node.ELEMENT_NODE) return;

  const el = child as Element;
  const tag = el.tagName.toLowerCase();
  if (STRIP_TAGS.has(tag) || isHiddenElement(el)) return;

  if (tag === 'br') {
    out.push(makeInline('\n', ctx));
    return;
  }
  if (tag === 'p' || tag === 'div') {
    appendWithBreak(out, inlineFromChildren(el, ctx));
    return;
  }

  const nextCtx = elementInlineCtx(el, ctx);
  out.push(...inlineFromChildren(el, nextCtx));
}

function elementInlineCtx(el: Element, ctx: InlineCtx): InlineCtx {
  const tag = el.tagName.toLowerCase();
  const next: InlineCtx = { ...ctx };
  const style = el.getAttribute('style') || '';
  const fs = /font-style\s*:\s*([^;]+)/i.exec(style)?.[1]?.trim().toLowerCase();

  if (tag === 'strong' || tag === 'b') next.bold = true;
  if (isStyleBold(el)) next.bold = true;
  if (tag === 'em' || tag === 'i') next.italic = true;
  if (fs === 'italic' || fs === 'oblique') next.italic = true;
  if (tag === 'code') next.code = true;
  if (tag === 'a') {
    const href = el.getAttribute('href');
    if (href) next.link = href;
  }
  return next;
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

function makeInline(text: string, ctx: InlineCtx): Inline {
  const inline: Inline = { text: decodeEntities(text) };
  if (ctx.bold) inline.bold = true;
  if (ctx.italic) inline.italic = true;
  if (ctx.code) inline.code = true;
  if (ctx.link) inline.link = ctx.link;
  return inline;
}

function appendWithBreak(target: Inline[], parts: Inline[]): void {
  if (!parts.length) return;
  if (target.length) target.push({ text: '\n' });
  target.push(...parts);
}
