import { describe, it, expect } from 'vitest';
import { detectFormat, formatLabel } from '../src/parse/detect';

describe('detectFormat', () => {
  it('detects markdown from a heading marker in the plain text', () => {
    expect(detectFormat({ text: '## Section\n\nSome text', html: '<p>## Section</p><p>Some text</p>' })).toBe('markdown');
  });

  it('detects markdown from bold markers in the plain text', () => {
    expect(detectFormat({ text: 'This is **bold** text.' })).toBe('markdown');
  });

  it('detects markdown from a pipe-table separator row', () => {
    const text = '| A | B |\n| --- | --- |\n| 1 | 2 |';
    expect(detectFormat({ text })).toBe('markdown');
  });

  it('does not mistake a plain horizontal rule for a table separator', () => {
    expect(detectFormat({ text: 'Before\n\n---\n\nAfter' })).toBe('markdown');
  });

  it('falls back to HTML when the HTML has real structure and the plain text has no markdown signals', () => {
    const html = '<h2>Heading</h2><p>Some <strong>bold</strong> text</p>';
    const text = 'Heading\nSome bold text';
    expect(detectFormat({ text, html })).toBe('html');
  });

  it('falls back to markdown (plain paragraphs) when neither signal is present', () => {
    expect(detectFormat({ text: 'Just a plain paragraph.', html: '<p>Just a plain paragraph.</p>' })).toBe('markdown');
    expect(detectFormat({ text: 'Just a plain paragraph, no html at all.' })).toBe('markdown');
  });

  it('prefers strong markdown signals over html structure', () => {
    const html = '<h2>Heading</h2>';
    const text = '## Heading';
    expect(detectFormat({ text, html })).toBe('markdown');
  });

  it('labels formats for display', () => {
    expect(formatLabel('markdown')).toBe('Markdown');
    expect(formatLabel('html')).toBe('Formatted text');
  });
});
