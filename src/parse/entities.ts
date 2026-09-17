/**
 * Minimal HTML entity decoder.
 *
 * marked's lexer intentionally leaves entities inside plain text tokens
 * un-decoded (that normally happens in its HTML renderer, which we bypass).
 * We only need to decode entities that plausibly show up in GPT-generated
 * Markdown, so this is a small named table plus numeric/hex references —
 * not a full HTML5 entity list.
 */
const NAMED: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
  mdash: '—',
  ndash: '–',
  hellip: '…',
  lsquo: '‘',
  rsquo: '’',
  ldquo: '“',
  rdquo: '”',
  copy: '©',
  reg: '®',
  trade: '™',
  euro: '€',
  deg: '°',
  times: '×'
};

export function decodeEntities(text: string): string {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match, body: string) => {
    if (body[0] === '#') {
      const isHex = body[1] === 'x' || body[1] === 'X';
      const code = parseInt(body.slice(isHex ? 2 : 1), isHex ? 16 : 10);
      if (Number.isFinite(code) && code > 0) {
        try {
          return String.fromCodePoint(code);
        } catch {
          return match;
        }
      }
      return match;
    }
    return Object.prototype.hasOwnProperty.call(NAMED, body) ? NAMED[body] : match;
  });
}
