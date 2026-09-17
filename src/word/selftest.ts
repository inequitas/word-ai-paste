/// <reference types="office-js" />
import { parseMarkdown } from '../parse/markdown';
import { normalize, DEFAULT_NORMALIZE_OPTIONS } from '../transform/normalize';
import { insertBlocks, DEFAULT_INSERT_OPTIONS } from './insert';
import { OfficeWordDocument, isWordApi13Supported } from './officeAdapter';

/**
 * Fixture exercised by the self-test: real H1-H3 headings, a bold pseudo
 * heading, mixed inline formatting, a citation marker, nested bullets, two
 * separate numbered lists (each must restart at 1), a table, and a quote —
 * one markdown string run through the exact same parse -> normalize ->
 * insert pipeline as a real paste.
 *
 * "Blank lines between blocks" is deliberately turned OFF for this fixture
 * (unlike the real default) so that a body paragraph and a heading sit
 * *directly* after a list with nothing normalize.ts would insert in
 * between — that adjacency is exactly what real Word can silently turn
 * into "just another list item" (see insert.ts's "Leaving a list cleanly"
 * doc comment), and unit tests can't catch that without a real Word
 * runtime. The blank-line spacing feature itself is already covered by
 * tests/normalize.test.ts. A genuine empty paragraph is then spliced in
 * by hand right after the third list, below, since Markdown itself has no
 * way to express an empty paragraph.
 */
const FIXTURE_MARKDOWN = `# Self-test heading 1

## Self-test heading 2

### Self-test heading 3

**Pseudo heading via bold line**

A paragraph with **bold**, *italic* and a [link](https://example.com) run, plus a citation marker【1†source】.

- Bullet one
  - Nested bullet
- Bullet two

Body text right after the bulleted list.

1. First numbered list, item one
2. First numbered list, item two

## A heading right after a numbered list

1. Second numbered list, item one
2. Second numbered list, item two

| Column A | Column B |
| --- | --- |
| a1 | b1 |
| a2 | b2 |

> A quoted line for the self-test.
`;

export interface SelfTestCheck {
  label: string;
  pass: boolean;
  detail?: string;
}

export interface SelfTestResult {
  ranAt: string;
  allPassed: boolean;
  checks: SelfTestCheck[];
}

export async function runSelfTest(): Promise<SelfTestResult> {
  if (!isWordApi13Supported()) {
    return {
      ranAt: new Date().toISOString(),
      allPassed: false,
      checks: [{ label: 'WordApi 1.3 available', pass: false, detail: 'This Word host is too old for AI Paste.' }]
    };
  }

  let blocks = normalize(parseMarkdown(FIXTURE_MARKDOWN), { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: false });
  // Splice a genuine empty paragraph in right after the third list (before
  // the table) — Markdown can't express one, but this is exactly the shape
  // of the "reused an empty bullet as the insertion point" / "blank line
  // right after a list" case the list-escape logic has to get right.
  const tableIndex = blocks.findIndex((b) => b.type === 'table');
  if (tableIndex !== -1) {
    blocks = [...blocks.slice(0, tableIndex), { type: 'paragraph', inlines: [] }, ...blocks.slice(tableIndex)];
  }
  const checks: SelfTestCheck[] = [];

  await Word.run(async (context) => {
    const startRange = context.document.getSelection().getRange(Word.RangeLocation.start);
    startRange.track();
    await context.sync();

    const doc = new OfficeWordDocument(context);
    await insertBlocks(doc, blocks, DEFAULT_INSERT_OPTIONS);

    const endRange = context.document.getSelection().getRange(Word.RangeLocation.start);
    const spanned = startRange.expandTo(endRange);
    const paragraphs = spanned.paragraphs;
    const tables = spanned.tables;
    paragraphs.load(
      'items/text,items/styleBuiltIn,items/listItemOrNullObject/level,items/listItemOrNullObject/listString'
    );
    tables.load('items/headerRowCount,items/rowCount');
    await context.sync();

    const texts = paragraphs.items.map((p) => p.text);
    const styles = paragraphs.items.map((p) => p.styleBuiltIn as string);
    const joinedText = texts.join('\n');

    checks.push({
      label: 'Heading 1, 2 and 3 use styleBuiltIn Heading1/Heading2/Heading3',
      pass: styles.includes('Heading1') && styles.includes('Heading2') && styles.includes('Heading3')
    });

    checks.push({
      label: 'Bold-only line was promoted to a heading',
      pass: paragraphs.items.some((p) => /pseudo heading/i.test(p.text) && /^Heading\d$/.test(p.styleBuiltIn as string))
    });

    checks.push({
      label: 'List paragraphs use the List Paragraph style (styleBuiltIn "ListParagraph")',
      pass: styles.includes('ListParagraph')
    });

    let nestedLevelSeen = false;
    const topLevelNumberedStarts: string[] = [];
    for (const p of paragraphs.items) {
      if (p.styleBuiltIn !== 'ListParagraph') continue;
      const item = p.listItemOrNullObject;
      if (item.isNullObject) continue;
      if (item.level > 0) nestedLevelSeen = true;
      if (item.level === 0 && /^\d+\.$/.test(item.listString.trim())) {
        topLevelNumberedStarts.push(item.listString.trim());
      }
    }
    checks.push({ label: 'A nested bullet level was created', pass: nestedLevelSeen });
    checks.push({
      label: 'Both separate numbered lists restart at "1."',
      pass: topLevelNumberedStarts.filter((s) => s === '1.').length >= 2,
      detail: topLevelNumberedStarts.join(', ') || '(none found)'
    });

    checks.push({
      label: 'A table with a header row was inserted',
      pass: tables.items.length > 0 && tables.items.some((t) => t.headerRowCount === 1 && t.rowCount >= 3)
    });

    checks.push({
      label: 'A Quote-styled paragraph was inserted',
      pass: styles.includes('Quote')
    });

    // The list-escape fix: none of these may come back as ListParagraph.
    const bodyAfterList = paragraphs.items.find((p) => /body text right after the bulleted list/i.test(p.text));
    checks.push({
      label: 'Body text right after a bulleted list is not a list item',
      pass: !!bodyAfterList && bodyAfterList.styleBuiltIn === 'Normal',
      detail: bodyAfterList ? `styleBuiltIn: ${bodyAfterList.styleBuiltIn}` : '(paragraph not found)'
    });

    const headingAfterList = paragraphs.items.find((p) => /heading right after a numbered list/i.test(p.text));
    checks.push({
      label: 'A heading right after a list is not a list item',
      pass: !!headingAfterList && /^Heading\d$/.test(headingAfterList.styleBuiltIn as string),
      detail: headingAfterList ? `styleBuiltIn: ${headingAfterList.styleBuiltIn}` : '(paragraph not found)'
    });

    const blankParagraphs = paragraphs.items.filter((p) => p.text.trim() === '');
    checks.push({
      label: 'A blank paragraph right after a list is not a list item',
      pass: blankParagraphs.length > 0 && blankParagraphs.every((p) => p.styleBuiltIn !== 'ListParagraph'),
      detail: blankParagraphs.length ? `${blankParagraphs.length} blank paragraph(s) found` : '(no blank paragraph found)'
    });

    checks.push({
      label: 'No leftover **, heading #, or 【…】 characters in the inserted text',
      pass: !joinedText.includes('**') && !/(^|\n)\s*#{1,6}\s/.test(joinedText) && !joinedText.includes('【')
    });

    startRange.untrack();
  });

  return {
    ranAt: new Date().toISOString(),
    allPassed: checks.every((c) => c.pass),
    checks
  };
}
