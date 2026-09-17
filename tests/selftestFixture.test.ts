import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { insertBlocks, DEFAULT_INSERT_OPTIONS } from '../src/word/insert';
import { buildFixtureBlocks, buildFixtureChecks, ooxmlBodyFacts, type SelfTestCheck } from '../src/word/selftestFixture';
import { FakeWordDocument } from './fakeWordAdapter';

const REQUIRED_LABELS = [
  'text order matches',
  'paragraph after each list is not a list item',
  'separate numbered lists restart at 1',
  "list indent matches Word's bullet button (≈36pt text / 18pt bullet)"
];

function byLabel(checks: SelfTestCheck[], label: string): SelfTestCheck {
  const found = checks.find((c) => c.label === label);
  if (!found) throw new Error(`no check labelled "${label}"`);
  return found;
}

beforeEach(() => {
  vi.spyOn(console, 'warn').mockImplementation(() => undefined);
});
afterEach(() => {
  vi.restoreAllMocks();
});

describe('self-test fixture', () => {
  it('has four list runs, including a bullet list directly followed by a numbered list', () => {
    const blocks = buildFixtureBlocks();
    const listIndexes = blocks.flatMap((b) => (b.type === 'listItem' ? [b.listIndex] : []));
    expect(new Set(listIndexes).size).toBe(4);
    const adjacent = blocks.some(
      (b, i) => b.type === 'listItem' && !b.ordered && blocks[i + 1]?.type === 'listItem' && (blocks[i + 1] as typeof b).ordered
    );
    expect(adjacent).toBe(true);
    expect(blocks.some((b) => b.type === 'listItem' && b.level === 1)).toBe(true);
    expect(blocks.some((b) => b.type === 'table')).toBe(true);
    // blank lines between blocks are on, as in real use
    expect(blocks.some((b) => b.type === 'paragraph' && b.inlines.length === 0)).toBe(true);
  });

  it('passes every check through the two-phase path on the fake adapter', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, buildFixtureBlocks(), DEFAULT_INSERT_OPTIONS);
    const checks = buildFixtureChecks(report);

    const failed = checks.filter((c) => !c.pass);
    expect(failed).toEqual([]);
    for (const label of REQUIRED_LABELS) expect(byLabel(checks, label).pass).toBe(true);
    for (const n of [1, 2, 3, 4]) {
      expect(byLabel(checks, `all items of list ${n} are list items`).pass).toBe(true);
      expect(byLabel(checks, `items of list ${n} share one list id`).pass).toBe(true);
    }
    expect(checks.every((c) => typeof c.detail === 'string' && c.detail.length > 0)).toBe(true);
    expect(byLabel(checks, 'separate numbered lists restart at 1').detail).toBe('list 3: 1. 2. 3. | list 4: 1. 2.');
  });

  it('explains the per-item fallback in the shared-list-id check, while numbering still restarts', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('attachToList', 99);
    const report = await insertBlocks(doc, buildFixtureBlocks(), DEFAULT_INSERT_OPTIONS);
    const checks = buildFixtureChecks(report);

    const shared = byLabel(checks, 'items of list 1 share one list id');
    expect(shared.pass).toBe(false);
    expect(shared.detail).toContain('per-item-lists fallback from item 2');
    expect(byLabel(checks, 'all items of list 1 are list items').pass).toBe(true);
    expect(byLabel(checks, 'separate numbered lists restart at 1').pass).toBe(true);
    expect(byLabel(checks, 'a nested list level was created (level 1)').pass).toBe(true);
  });

  it('recovers with direct paragraph indents when setLevelIndents could not be set', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('setLevelIndents', 99);
    const report = await insertBlocks(doc, buildFixtureBlocks(), DEFAULT_INSERT_OPTIONS);
    const check = byLabel(buildFixtureChecks(report), "list indent matches Word's bullet button (≈36pt text / 18pt bullet)");
    // setIndents is called as a fallback, so the indents should still be correct
    expect(check.pass).toBe(true);
  });

  it('names the failing items when a list item did not become one', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.attachToListIsNoop = true;
    doc.failNext('startNewList', 1, 1); // list 1's first item works; its item 2 can't get its own list either
    const report = await insertBlocks(doc, buildFixtureBlocks(), DEFAULT_INSERT_OPTIONS);
    const check = byLabel(buildFixtureChecks(report), 'all items of list 1 are list items');
    expect(check.pass).toBe(false);
    expect(check.detail).toMatch(/^not list items: item 2 "Alpha item two"/);
  });

  it('fails every required check with a reason when there is no read-back', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    doc.failNext('readBack');
    const report = await insertBlocks(doc, buildFixtureBlocks(), DEFAULT_INSERT_OPTIONS);
    const checks = buildFixtureChecks(report);
    expect(byLabel(checks, 'read-back of the inserted range succeeded').detail).toContain('GeneralException');
    for (const label of [...REQUIRED_LABELS, 'items of list 2 share one list id']) {
      expect(byLabel(checks, label)).toMatchObject({ pass: false, detail: 'skipped: no read-back' });
    }
  });

  it('reports a text-order problem with the paragraph position', async () => {
    const doc = new FakeWordDocument({ anchorText: '' });
    const report = await insertBlocks(doc, buildFixtureBlocks(), DEFAULT_INSERT_OPTIONS);
    const paragraphs = [...report.readBack!.paragraphs];
    [paragraphs[5], paragraphs[6]] = [paragraphs[6], paragraphs[5]];
    const check = byLabel(buildFixtureChecks({ ...report, readBack: { ...report.readBack!, paragraphs } }), 'text order matches');
    expect(check.pass).toBe(false);
    expect(check.detail).toBe('paragraph #6: expected "Alpha nested item", found "Alpha item three"');
  });
});

describe('ooxmlBodyFacts', () => {
  const pkg = (body: string, extra = ''): string =>
    `<pkg:package><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body>${body}</w:body></w:document></pkg:xmlData></pkg:part>${extra}</pkg:package>`;

  it('finds numPr and pStyle inside the document body', () => {
    const xml = pkg('<w:p><w:pPr><w:pStyle w:val="Lijstalinea"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr></w:p>');
    expect(ooxmlBodyFacts(xml)).toEqual({ numPr: true, pStyle: 'Lijstalinea' });
  });

  it('ignores numPr/pStyle that only appear in other parts (styles, numbering)', () => {
    const xml = pkg(
      '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      '<pkg:part pkg:name="/word/numbering.xml"><w:lvl><w:pStyle w:val="Heading1"/></w:lvl></pkg:part><w:numPr/>'
    );
    expect(ooxmlBodyFacts(xml)).toEqual({ numPr: false, pStyle: null });
  });
});
