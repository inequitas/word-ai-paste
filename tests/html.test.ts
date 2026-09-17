import { describe, it, expect } from 'vitest';
import { parseHtml } from '../src/parse/html';
import { plainText } from '../src/model';
import type { HeadingBlock, ListItemBlock, TableBlock, CodeBlock, QuoteBlock, ParagraphBlock } from '../src/model';

describe('parseHtml — generic chat-tool HTML', () => {
  it('parses headings, bold/italic/link and drops class/style noise', () => {
    const html = `
      <div class="markdown prose">
        <h2 class="fancy-heading" style="color:red">Section heading</h2>
        <p>Some <strong>bold</strong> and <em>italic</em> and a
        <a href="https://example.com">link</a> here.</p>
      </div>`;
    const blocks = parseHtml(html);
    expect(blocks[0].type).toBe('heading');
    expect((blocks[0] as HeadingBlock).level).toBe(2);
    expect(plainText((blocks[0] as HeadingBlock).inlines)).toBe('Section heading');

    const p = blocks[1] as ParagraphBlock;
    expect(p.inlines.find((i) => i.text === 'bold')?.bold).toBe(true);
    expect(p.inlines.find((i) => i.text === 'italic')?.italic).toBe(true);
    expect(p.inlines.find((i) => i.text === 'link')?.link).toBe('https://example.com');
  });

  it('detects bold via inline font-weight style, not just <b>/<strong>', () => {
    const html = '<p><span style="font-weight:700">Strong-ish</span> and <span style="font-weight:400">normal</span></p>';
    const [p] = parseHtml(html) as [ParagraphBlock];
    expect(p.inlines.find((i) => i.text === 'Strong-ish')?.bold).toBe(true);
    expect(p.inlines.find((i) => i.text === 'normal')?.bold).toBeUndefined();
  });

  it('detects italic via inline font-style style', () => {
    const html = '<p><span style="font-style:italic">Slanted</span></p>';
    const [p] = parseHtml(html) as [ParagraphBlock];
    expect(p.inlines[0].italic).toBe(true);
  });

  it('handles nested <ul> with <p> inside <li>', () => {
    const html = `
      <ul>
        <li><p>item 1</p><ul><li>nested item</li></ul></li>
        <li>item 2</li>
      </ul>`;
    const items = parseHtml(html) as ListItemBlock[];
    expect(items).toHaveLength(3);
    expect(items[0]).toMatchObject({ level: 0, ordered: false });
    expect(plainText(items[0].inlines)).toBe('item 1');
    expect(items[1]).toMatchObject({ level: 1, ordered: false });
    expect(items[2]).toMatchObject({ level: 0, ordered: false });
  });

  it('gives two separate <ol> elements different listIndex values', () => {
    const html = '<ol><li>a1</li><li>a2</li></ol><p>between</p><ol><li>b1</li><li>b2</li></ol>';
    const items = parseHtml(html).filter((b): b is ListItemBlock => b.type === 'listItem');
    expect(items[0].listIndex).toBe(items[1].listIndex);
    expect(items[2].listIndex).toBe(items[3].listIndex);
    expect(items[0].listIndex).not.toBe(items[2].listIndex);
  });

  it('parses a table with thead/tbody', () => {
    const html = `
      <table>
        <thead><tr><th>A</th><th>B</th></tr></thead>
        <tbody><tr><td>1</td><td>2</td></tr></tbody>
      </table>`;
    const [table] = parseHtml(html) as [TableBlock];
    expect(table.header).not.toBeNull();
    expect(plainText(table.header![0])).toBe('A');
    expect(table.rows).toHaveLength(1);
  });

  it('parses a blockquote and a <pre><code> block', () => {
    const html = '<blockquote><p>quoted text</p></blockquote><pre><code>const x = 1;\nconsole.log(x);</code></pre>';
    const blocks = parseHtml(html);
    expect((blocks[0] as QuoteBlock).type).toBe('quote');
    expect(plainText((blocks[0] as QuoteBlock).inlines)).toBe('quoted text');
    expect((blocks[1] as CodeBlock).text).toBe('const x = 1;\nconsole.log(x);');
  });

  it('strips a chat-tool code toolbar (language label + Copy code button) before a <pre>', () => {
    const html = `
      <div class="code-block-wrapper">
        <div class="code-toolbar">
          <span>javascript</span>
          <button>Copy code</button>
        </div>
        <pre><code class="language-javascript">const x = 1;</code></pre>
      </div>`;
    const blocks = parseHtml(html);
    expect(blocks).toHaveLength(1);
    expect(blocks[0].type).toBe('code');
    expect((blocks[0] as CodeBlock).text).toBe('const x = 1;');
  });

  it('strips script/style/svg/button and screen-reader-only content', () => {
    const html = `
      <p>Visible text
        <span class="sr-only">hidden helper text</span>
        <svg><circle /></svg>
        <button>Click</button>
        <script>var scriptRanButDoesNothingHarmful = 1;</script>
      </p>`;
    const [p] = parseHtml(html) as [ParagraphBlock];
    // Whitespace trimming/collapsing is normalize.ts's job, not the parser's.
    expect(plainText(p.inlines).trim()).toBe('Visible text');
  });

  it('treats a stray <div> paragraph like a <p>', () => {
    const [p] = parseHtml('<div>Just a div paragraph</div>') as [ParagraphBlock];
    expect(p.type).toBe('paragraph');
    expect(plainText(p.inlines)).toBe('Just a div paragraph');
  });

  it('turns <br> into a "\\n" inline', () => {
    const [p] = parseHtml('<p>Line one<br>Line two</p>') as [ParagraphBlock];
    expect(p.inlines.some((i) => i.text === '\n')).toBe(true);
  });
});

describe('parseHtml — Word ("Mso*") HTML export', () => {
  it('recognizes MsoHeadingN classes as headings', () => {
    const html = '<p class="MsoHeading1">Chapter title</p><p class="MsoHeading2">Section title</p>';
    const blocks = parseHtml(html) as HeadingBlock[];
    expect(blocks[0].level).toBe(1);
    expect(blocks[1].level).toBe(2);
    expect(plainText(blocks[0].inlines)).toBe('Chapter title');
  });

  it('recognizes MsoListParagraph + mso-list style as bulleted list items and strips the marker span', () => {
    const html =
      '<p class="MsoListParagraph" style="mso-list:l0 level1 lfo1">' +
      '<span style="mso-list:Ignore">&middot;<span>&nbsp;&nbsp;</span></span>' +
      'First bullet</p>' +
      '<p class="MsoListParagraph" style="mso-list:l0 level1 lfo1">' +
      '<span style="mso-list:Ignore">&middot;<span>&nbsp;&nbsp;</span></span>' +
      'Second bullet</p>';
    const items = parseHtml(html) as ListItemBlock[];
    expect(items).toHaveLength(2);
    expect(items[0].ordered).toBe(false);
    expect(items[0].level).toBe(0);
    expect(items[0].listIndex).toBe(items[1].listIndex);
    expect(plainText(items[0].inlines)).toBe('First bullet');
  });

  it('recognizes a numbered mso-list paragraph via its marker text', () => {
    const html =
      '<p class="MsoListParagraph" style="mso-list:l1 level1 lfo2">' +
      '<span style="mso-list:Ignore">1.<span>&nbsp;</span></span>' +
      'First numbered</p>';
    const [item] = parseHtml(html) as ListItemBlock[];
    expect(item.ordered).toBe(true);
    expect(plainText(item.inlines)).toBe('First numbered');
  });

  it('gives different mso-list ids separate listIndex values', () => {
    const html =
      '<p class="MsoListParagraph" style="mso-list:l0 level1 lfo1"><span style="mso-list:Ignore">&middot;</span>a</p>' +
      '<p class="MsoListParagraph" style="mso-list:l3 level1 lfo2"><span style="mso-list:Ignore">&middot;</span>b</p>';
    const items = parseHtml(html) as ListItemBlock[];
    expect(items[0].listIndex).not.toBe(items[1].listIndex);
  });
});

describe('parseHtml — Cocoa HTML Writer output (macOS RTF -> HTML, e.g. textutil)', () => {
  const cocoaHtml = `
    <html>
    <head>
      <style type="text/css">
        p.p1 {margin: 0.0px; font: 24.0px Times}
        p.p2 {margin: 0.0px; font: 18.0px Times}
        p.p3 {margin: 0.0px; font: 12.0px Times}
        p.p5 {margin: 0.0px; text-align: center; font: 12.0px Times}
        span.s1 {font-kerning: none}
      </style>
    </head>
    <body>
      <p class="p1"><span class="s1"><b>Scope and assumptions</b></span></p>
      <p class="p2"><span class="s1"><b>In scope</b></span></p>
      <p class="p3"><span class="s1">The project covers the design and rollout of a new network access solution. Key tasks:</span></p>
      <ul class="ul1"><li class="li4"><span class="s1">first work package;</span></li></ul>
      <table cellspacing="0" cellpadding="0">
        <tbody>
          <tr>
            <td valign="middle"><p class="p5"><span class="s1"><b>Location</b></span></p></td>
            <td valign="middle"><p class="p5"><span class="s1"><b>Work</b></span></p></td>
          </tr>
          <tr>
            <td valign="middle"><p class="p3"><span class="s1">Site A</span></p></td>
            <td valign="middle"><p class="p3"><span class="s1">Pilot rollout</span></p></td>
          </tr>
        </tbody>
      </table>
    </body>
    </html>`;

  it('treats large bold paragraphs (per the <style> block sizing) as ranked headings, body text as plain paragraphs', () => {
    const blocks = parseHtml(cocoaHtml);
    expect(blocks[0].type).toBe('heading');
    expect((blocks[0] as HeadingBlock).level).toBe(1);
    expect(plainText((blocks[0] as HeadingBlock).inlines)).toBe('Scope and assumptions');

    expect(blocks[1].type).toBe('heading');
    expect((blocks[1] as HeadingBlock).level).toBe(2);
    expect(plainText((blocks[1] as HeadingBlock).inlines)).toBe('In scope');

    expect(blocks[2].type).toBe('paragraph');
    expect(plainText((blocks[2] as ParagraphBlock).inlines)).toContain('The project covers');
  });

  it('treats a table\'s first row as a header when every cell is bold, with no <th> present', () => {
    const [table] = parseHtml(cocoaHtml).filter((b): b is TableBlock => b.type === 'table');
    expect(table.header).not.toBeNull();
    expect(plainText(table.header![0])).toBe('Location');
    expect(table.rows).toHaveLength(1);
    expect(plainText(table.rows[0][0])).toBe('Site A');
  });

  it('does not treat a same-size bold body line as a heading (real <h1-6> / size-ranked bold only)', () => {
    const html = `
      <style>p.p3 {font: 12.0px Times}</style>
      <p class="p3"><b>Bold but body-sized</b></p>
      <p class="p3">Body text making body the most common size.</p>
      <p class="p3">More body text so body size wins by character count.</p>`;
    const blocks = parseHtml(html);
    expect(blocks.every((b) => b.type === 'paragraph')).toBe(true);
  });
});

describe('parseHtml — Dutch text', () => {
  it('parses a Dutch heading and a bold label paragraph from chat-tool HTML', () => {
    const html = '<h2>1. Inleiding</h2><p><strong>Doelstelling:</strong> dit project heeft als doel de netwerktoegang te vernieuwen.</p>';
    const blocks = parseHtml(html);
    expect(plainText((blocks[0] as HeadingBlock).inlines)).toBe('1. Inleiding');
    const p = blocks[1] as ParagraphBlock;
    expect(p.inlines[0].text).toBe('Doelstelling:');
    expect(p.inlines[0].bold).toBe(true);
  });
});
