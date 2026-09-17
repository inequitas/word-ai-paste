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
 */
const FIXTURE_MARKDOWN = `# Self-test heading 1

## Self-test heading 2

### Self-test heading 3

**Pseudo heading via bold line**

A paragraph with **bold**, *italic* and a [link](https://example.com) run, plus a citation marker【1†source】.

- Bullet one
  - Nested bullet
- Bullet two

1. First numbered list, item one
2. First numbered list, item two

Some text between the two numbered lists.

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

  const blocks = normalize(parseMarkdown(FIXTURE_MARKDOWN), DEFAULT_NORMALIZE_OPTIONS);
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
