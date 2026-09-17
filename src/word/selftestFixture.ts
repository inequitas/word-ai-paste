import type { Block } from '../model';
import { parseMarkdown } from '../parse/markdown';
import { normalize, DEFAULT_NORMALIZE_OPTIONS } from '../transform/normalize';
import type { InsertReport } from './insert';
import type { ReadBackParagraph } from './adapter';
import { firstTextDifference, indentProblem, listLevelIndents, normalizeText, round, runItems, runListIds, snippet } from './verify';

/**
 * The self-test fixture and its checks — pure (no Office.js), so the whole
 * fixture can also be run against the fake adapter in the unit tests.
 *
 * The fixture mirrors the shape that broke in real Word for Mac with
 * v1.0.1 (heading, body ending in a colon, a 6-item bullet list, heading,
 * body, body, another list), with the real default options (blank lines
 * between blocks ON), plus: a nested bullet, a promoted bold line, mixed
 * inline formatting and a citation marker, a bullet list directly followed
 * by a numbered list, two separate numbered lists (each must restart at
 * 1), a table and a quote. All text is synthetic.
 */
export const FIXTURE_MARKDOWN = `# Self-test heading 1

## Self-test heading 2

Body text before the first list, ending with a colon:

- Alpha item one
- Alpha item two
  - Alpha nested item
- Alpha item three
- Alpha item four
- Alpha item five
- Alpha item six

### Self-test heading 3

**Pseudo heading via bold line**

A paragraph with **bold**, *italic* and a [link](https://example.com) run, plus a citation marker【1†source】.

A second body paragraph, ending with a colon:

- Beta item one
- Beta item two
- Beta item three

1. First numbered list, item one
2. First numbered list, item two
3. First numbered list, item three

Body text between the two numbered lists.

1. Second numbered list, item one
2. Second numbered list, item two

| Column A | Column B |
| --- | --- |
| a1 | b1 |
| a2 | b2 |

> A quoted line for the self-test.
`;

export function buildFixtureBlocks(): Block[] {
  return normalize(parseMarkdown(FIXTURE_MARKDOWN), { ...DEFAULT_NORMALIZE_OPTIONS, blankLinesBetweenBlocks: true });
}

export interface SelfTestCheck {
  label: string;
  pass: boolean;
  detail?: string;
  /** Informational probe: shown, but never counted as a failure unless it threw. */
  info?: boolean;
}

/** Checks on the fixture insert, from its InsertReport (read-back included). */
export function buildFixtureChecks(report: InsertReport): SelfTestCheck[] {
  const checks: SelfTestCheck[] = [];
  const rb = report.readBack;
  const runs = report.runs;
  const listLabels = runs.flatMap((r) => [`all items of list ${r.run} are list items`, `items of list ${r.run} share one list id`]);

  checks.push({
    label: 'read-back of the inserted range succeeded',
    pass: rb !== null,
    detail: rb
      ? `${rb.paragraphs.length} paragraphs, ${rb.tables.length} table(s)${rb.notes.length ? `; ${rb.notes.join('; ')}` : ''}`
      : `read-back failed: ${report.verifyError ?? 'no paragraphs inserted'}`
  });

  if (!rb) {
    const skipped = [
      'text order matches',
      ...listLabels,
      'paragraph after each list is not a list item',
      'separate numbered lists restart at 1',
      "list indent matches Word's bullet button (≈36pt text / 18pt bullet)"
    ];
    for (const label of skipped) checks.push({ label, pass: false, detail: 'skipped: no read-back' });
    return checks;
  }

  const actual = rb.paragraphs;
  const expected = report.expected;

  // Text order
  const diff = firstTextDifference(expected, actual);
  checks.push({
    label: 'text order matches',
    pass: diff === null,
    detail: diff ?? `${expected.length} paragraphs in the expected order`
  });

  // Per list run: membership and one shared list id
  for (const run of runs) {
    const items = runItems(run, actual);
    const kind = `${run.ordered ? 'numbered' : 'bulleted'}, ${run.itemCount} items`;
    const mode = run.mode === 'per-item' ? `per-item-lists fallback from item ${(run.fallbackFrom ?? 0) + 1}` : 'one list';
    const notItems = items
      .map((p, k) => ({ p, k }))
      .filter(({ p }) => !p || !p.isListItem)
      .map(({ p, k }) => `item ${k + 1} ${p ? snippet(p.text, 25) : '(missing)'}`);
    checks.push({
      label: `all items of list ${run.run} are list items`,
      pass: notItems.length === 0,
      detail: notItems.length ? `not list items: ${notItems.join(', ')} (${kind}, ${mode})` : `${kind}, ${mode}`
    });

    const ids = runListIds(run, actual);
    const distinct = new Set(ids);
    const shared = ids.length > 0 && !distinct.has(null) && distinct.size === 1;
    checks.push({
      label: `items of list ${run.run} share one list id`,
      pass: shared,
      detail:
        `ids: ${ids.map((id) => id ?? '-').join(', ')}` +
        (run.mode === 'per-item' ? ` (${mode})` : '') +
        (ids.every((id) => id === null) && rb.notes.length ? ` (${rb.notes.join('; ')})` : '')
    });
  }

  // The paragraph right after each list
  const afterProblems: string[] = [];
  const afterChecked: string[] = [];
  for (const run of runs) {
    const idx = run.firstExpected + run.itemCount;
    const e = expected[idx];
    if (!e || e.listItem) continue; // end of the insert, or directly followed by another list
    const a = actual[idx];
    if (!a) {
      afterProblems.push(`list ${run.run}: paragraph #${idx + 1} missing`);
    } else if (a.isListItem || normalizeText(a.text) !== normalizeText(e.text)) {
      afterProblems.push(
        `list ${run.run}: next paragraph ${snippet(a.text, 25)} ${a.isListItem ? 'is a list item' : `is not the expected ${snippet(e.text, 25)}`}`
      );
    } else {
      afterChecked.push(`list ${run.run} → ${snippet(a.text, 20)} (${a.styleBuiltIn})`);
    }
  }
  checks.push({
    label: 'paragraph after each list is not a list item',
    pass: afterProblems.length === 0 && afterChecked.length > 0,
    detail: afterProblems.length ? afterProblems.join('; ') : afterChecked.join('; ') || 'no list is followed by a paragraph'
  });

  // Numbered lists restart at 1
  const orderedRuns = runs.filter((r) => r.ordered);
  const sequences = orderedRuns.map((run) => {
    const items = runItems(run, actual);
    const top = items.filter((p, k) => p && expected[run.firstExpected + k]?.level === 0) as ReadBackParagraph[];
    const got = top.map((p) => (p.listString ?? '?').trim());
    const want = top.map((_, k) => `${k + 1}.`);
    return { run, got, ok: got.length > 0 && got.every((s, k) => s === want[k]) };
  });
  checks.push({
    label: 'separate numbered lists restart at 1',
    pass: sequences.length >= 2 && sequences.every((s) => s.ok),
    detail: sequences.length
      ? sequences.map((s) => `list ${s.run.run}: ${s.got.join(' ') || '(none)'}${s.ok ? '' : ' ✗'}`).join(' | ')
      : 'no numbered lists found'
  });

  // Indents
  const indentIssues: string[] = [];
  const seenLevels = new Map<number, string>();
  expected.forEach((e, i) => {
    if (!e.listItem) return;
    const a = actual[i];
    if (!a || !a.isListItem) return;
    const problem = indentProblem(a, e.level);
    if (problem) indentIssues.push(`item ${snippet(a.text, 20)}: ${problem}`);
    if (!seenLevels.has(e.level) && a.leftIndent !== null && a.firstLineIndent !== null) {
      seenLevels.set(
        e.level,
        `level ${e.level}: text ${round(a.leftIndent)}pt, bullet ${round(a.leftIndent + a.firstLineIndent)}pt (want ${
          listLevelIndents(e.level).textIndent
        }/${listLevelIndents(e.level).bulletIndent})`
      );
    }
  });
  checks.push({
    label: "list indent matches Word's bullet button (≈36pt text / 18pt bullet)",
    pass: indentIssues.length === 0 && seenLevels.size > 0,
    detail: indentIssues.length
      ? `${indentIssues.length} item(s) off: ${indentIssues.slice(0, 3).join('; ')}${indentIssues.length > 3 ? '; …' : ''}`
      : Array.from(seenLevels.values()).join('; ') || 'no list items read back'
  });

  // Styles and content
  const styles = actual.map((p) => p.styleBuiltIn);
  const headingStyles = ['Heading1', 'Heading2', 'Heading3'];
  const missingHeadings = headingStyles.filter((s) => !styles.includes(s));
  checks.push({
    label: 'Heading 1, 2 and 3 use styleBuiltIn Heading1/Heading2/Heading3',
    pass: missingHeadings.length === 0,
    detail: missingHeadings.length ? `missing: ${missingHeadings.join(', ')}` : 'all three found'
  });

  const pseudo = actual.find((p) => /pseudo heading/i.test(p.text));
  checks.push({
    label: 'bold-only line was promoted to a heading',
    pass: !!pseudo && /^Heading\d$/.test(pseudo.styleBuiltIn),
    detail: pseudo ? `styleBuiltIn: ${pseudo.styleBuiltIn}` : '(paragraph not found)'
  });

  const listItemStyles = expected
    .map((e, i) => ({ e, a: actual[i] }))
    .filter(({ e }) => e.listItem)
    .filter(({ a }) => !a || a.styleBuiltIn !== 'ListParagraph')
    .map(({ e, a }) => `${snippet(e.text, 20)}: ${a ? a.styleBuiltIn : 'missing'}`);
  checks.push({
    label: 'list items use the List Paragraph style',
    pass: listItemStyles.length === 0,
    detail: listItemStyles.length ? listItemStyles.join('; ') : 'all list items are ListParagraph'
  });

  const nested = actual.filter((p) => p.isListItem && p.level === 1);
  checks.push({
    label: 'a nested list level was created (level 1)',
    pass: nested.length > 0,
    detail: nested.length
      ? `${nested.length} item(s) at level 1`
      : `levels found: ${Array.from(new Set(actual.filter((p) => p.isListItem).map((p) => p.level ?? '?'))).join(', ') || 'none'}`
  });

  checks.push({
    label: 'a table with a header row was inserted',
    pass: rb.tables.some((t) => t.headerRowCount === 1 && t.rowCount >= 3),
    detail: rb.tables.length
      ? rb.tables.map((t) => `${t.rowCount} rows, headerRowCount ${t.headerRowCount}`).join('; ')
      : 'no table in the read-back range'
  });

  checks.push({
    label: 'a Quote-styled paragraph was inserted',
    pass: styles.includes('Quote'),
    detail: styles.includes('Quote') ? 'found' : 'no paragraph with styleBuiltIn Quote'
  });

  const styleMismatches = report.mismatches.filter((m) => m.includes('style should be'));
  checks.push({
    label: 'every paragraph has its intended style (headings, body, blank lines)',
    pass: styleMismatches.length === 0,
    detail: styleMismatches.length ? styleMismatches.slice(0, 3).join('; ') : 'all styles as intended'
  });

  const joined = actual.map((p) => normalizeText(p.text)).join('\n');
  const leftovers = [
    joined.includes('**') ? '**' : null,
    /(^|\n)\s*#{1,6}\s/.test(joined) ? 'heading #' : null,
    joined.includes('【') ? '【…】' : null
  ].filter(Boolean);
  checks.push({
    label: 'no leftover **, heading #, or 【…】 characters in the inserted text',
    pass: leftovers.length === 0,
    detail: leftovers.length ? `found: ${leftovers.join(', ')}` : 'clean'
  });

  checks.push({
    label: 'read-back comparison found no mismatches',
    pass: report.mismatches.length === 0,
    detail: report.mismatches.length
      ? `${report.mismatches.length}: ${report.mismatches.slice(0, 3).join('; ')}${report.mismatches.length > 3 ? '; …' : ''}`
      : 'none'
  });

  return checks;
}

/** What a paragraph's getOoxml() says about list numbering and style, looking only at the document body part. */
export function ooxmlBodyFacts(xml: string): { numPr: boolean; pStyle: string | null } {
  const m = xml.match(/<w:body\b[^>]*>([\s\S]*?)<\/w:body>/);
  const body = m ? m[1] : xml;
  const style = body.match(/<w:pStyle\b[^>]*w:val="([^"]*)"/);
  return { numPr: /<w:numPr\b/.test(body), pStyle: style ? style[1] : null };
}
