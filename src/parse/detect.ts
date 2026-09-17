export type DetectedFormat = 'markdown' | 'html';

export interface DetectInput {
  text: string;
  html?: string;
}

/**
 * Decides whether a paste should be parsed as Markdown or as HTML.
 *
 * Strong Markdown signals in the plain-text flavor win outright (a GPT
 * chat's Markdown reply, once it round-trips through the OS clipboard,
 * often also carries an HTML flavor that's just `<p>` wrapping the same
 * literal `##`/`**`/`|` characters — that HTML has "structure" too, but
 * it's still Markdown). Otherwise, real HTML structure in the HTML flavor
 * wins. With neither, it's plain paragraphs, handled as Markdown (a no-op
 * parse).
 */
export function detectFormat(input: DetectInput): DetectedFormat {
  const text = input.text ?? '';
  if (hasStrongMarkdownSignals(text)) return 'markdown';
  if (input.html && hasHtmlStructure(input.html)) return 'html';
  return 'markdown';
}

export function formatLabel(format: DetectedFormat): string {
  return format === 'markdown' ? 'Markdown' : 'Formatted text';
}

function hasStrongMarkdownSignals(text: string): boolean {
  const lines = text.split('\n');
  if (lines.some((l) => /^#{1,6}\s+\S/.test(l.trim()))) return true;
  if (/\*\*[^*\n]+\*\*/.test(text)) return true;
  if (hasMarkdownTable(lines)) return true;
  return false;
}

function hasMarkdownTable(lines: string[]): boolean {
  for (let i = 1; i < lines.length; i++) {
    const row = lines[i - 1].trim();
    const sep = lines[i].trim();
    if (row.includes('|') && isTableSeparatorLine(sep)) return true;
  }
  return false;
}

function isTableSeparatorLine(line: string): boolean {
  if (!line.includes('|') || !line.includes('-')) return false;
  return /^[\s|:-]+$/.test(line);
}

function hasHtmlStructure(html: string): boolean {
  return /<(h[1-6]|ul|ol|table|strong|b|em)\b/i.test(html);
}
