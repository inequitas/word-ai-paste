/// <reference types="office-js" />
import { DEFAULT_INSERT_OPTIONS, InsertError, describeError, formatStep, insertBlocks } from './insert';
import { OfficeWordDocument, isWordApi13Supported } from './officeAdapter';
import { buildListOoxml } from './ooxml';
import { buildFixtureBlocks, buildFixtureChecks, ooxmlBodyFacts, type SelfTestCheck } from './selftestFixture';
import { listLevelIndents, round } from './verify';

export type { SelfTestCheck } from './selftestFixture';

export interface SelfTestResult {
  ranAt: string;
  allPassed: boolean;
  /** Small isolated probes of individual Office.js calls, run first — see runCapabilityProbes. */
  probes: SelfTestCheck[];
  /** Checks on the full markdown -> normalize -> two-phase insert fixture. */
  checks: SelfTestCheck[];
  /** The fixture insert's step log. */
  steps: string[];
  /** The fixture insert's read-back mismatches. */
  mismatches: string[];
  /** The step that was in progress when something threw, if anything did. */
  failedStep: string | null;
}

/** Counts toward "all passed": everything except informational probes that did not throw. */
function counts(c: SelfTestCheck): boolean {
  return !c.info || !c.pass;
}

export async function runSelfTest(): Promise<SelfTestResult> {
  const result: SelfTestResult = {
    ranAt: new Date().toISOString(),
    allPassed: false,
    probes: [],
    checks: [],
    steps: [],
    mismatches: [],
    failedStep: null
  };

  if (!isWordApi13Supported()) {
    result.checks.push({ label: 'WordApi 1.3 available', pass: false, detail: 'This Word host is too old for AI Paste.' });
    return result;
  }

  const blocks = buildFixtureBlocks();
  let stage = 'starting Word.run';

  try {
    await Word.run(async (context) => {
      stage = 'capability probes';
      result.probes = await runCapabilityProbes(context, (label) => {
        stage = `probe "${label}"`;
      });

      stage = 'fixture insert';
      let report;
      try {
        report = await insertBlocks(new OfficeWordDocument(context), blocks, DEFAULT_INSERT_OPTIONS);
      } catch (err) {
        if (!(err instanceof InsertError)) throw err;
        result.steps = err.steps.map(formatStep);
        result.failedStep = err.failedStep ? formatStep(err.failedStep) : 'setup';
        result.checks.push({
          label: 'fixture insert completed',
          pass: false,
          detail: `failed at ${result.failedStep}: ${describeError(err.rawError)}`
        });
        return;
      }

      result.steps = report.steps.map(formatStep);
      result.mismatches = report.mismatches;
      const listSummary = report.runs
        .map((r) => `list ${r.run}: ${r.mode === 'list' ? 'one list' : `per-item from item ${(r.fallbackFrom ?? 0) + 1}`}`)
        .join(', ');
      result.checks.push({
        label: 'fixture insert completed',
        pass: true,
        detail: `${blocks.length} blocks, ${report.expected.length} paragraphs; ${listSummary || 'no lists'}${
          report.repaired ? `; repair pass detached ${report.repaired}` : ''
        }`
      });

      stage = 'fixture checks';
      result.checks.push(...buildFixtureChecks(report));
    });
  } catch (err) {
    result.failedStep = result.failedStep ?? stage;
    result.checks.push({ label: 'self-test ran to the end', pass: false, detail: `threw during ${stage}: ${describeError(err)}` });
  }

  result.allPassed = result.checks.length > 0 && [...result.probes, ...result.checks].filter(counts).every((c) => c.pass);
  return result;
}

// ---------------------------------------------------------------------------
// Capability probes — small, isolated checks of individual Office.js calls,
// run at the end of the document. Each probe is its own try + sync (so a
// GeneralException is attributed to exactly one call, with debugInfo),
// reports the values it actually read back, and deletes what it created.
// ---------------------------------------------------------------------------

interface ProbeScope {
  /** A new Normal paragraph at the end of the document that is NOT a list item (inherited membership is detached and noted). */
  plainParagraph(text: string): Promise<Word.Paragraph>;
  /** Registers something to delete after the probe. */
  created(obj: { delete(): void }): void;
}

interface ProbeOutcome {
  pass: boolean;
  detail: string;
}

interface ParaFacts {
  isListItem: boolean;
  styleBuiltIn: string;
  leftIndent: number;
  firstLineIndent: number;
  listId: number | null;
  level: number | null;
  listString: string | null;
}

async function readFacts(context: Word.RequestContext, p: Word.Paragraph): Promise<ParaFacts> {
  p.load('isListItem,styleBuiltIn,leftIndent,firstLineIndent');
  await context.sync();
  let listId: number | null = null;
  let level: number | null = null;
  let listString: string | null = null;
  if (p.isListItem) {
    const list = p.listOrNullObject;
    const item = p.listItemOrNullObject;
    list.load('id');
    item.load('level,listString');
    await context.sync();
    listId = list.isNullObject ? null : list.id;
    level = item.isNullObject ? null : item.level;
    listString = item.isNullObject ? null : item.listString;
  }
  return {
    isListItem: p.isListItem,
    styleBuiltIn: String(p.styleBuiltIn),
    leftIndent: p.leftIndent,
    firstLineIndent: p.firstLineIndent,
    listId,
    level,
    listString
  };
}

function describeFacts(f: ParaFacts): string {
  const parts = [`isListItem=${f.isListItem}`, `style=${f.styleBuiltIn}`];
  if (f.isListItem) parts.push(`listId=${f.listId ?? '?'}`, `level=${f.level ?? '?'}`, `listString=${JSON.stringify(f.listString)}`);
  parts.push(indentFacts(f));
  return parts.join(', ');
}

function indentFacts(f: ParaFacts): string {
  return `leftIndent=${round(f.leftIndent)}, firstLineIndent=${round(f.firstLineIndent)} (text at ${round(f.leftIndent)}pt, bullet/first line at ${round(
    f.leftIndent + f.firstLineIndent
  )}pt)`;
}

async function runProbe(
  context: Word.RequestContext,
  label: string,
  options: { info?: boolean },
  body: (scope: ProbeScope) => Promise<ProbeOutcome>
): Promise<SelfTestCheck> {
  const created: { delete(): void }[] = [];
  const notes: string[] = [];
  const scope: ProbeScope = {
    created: (obj) => created.push(obj),
    plainParagraph: async (text) => {
      const p = context.document.body.insertParagraph(text, Word.InsertLocation.end);
      created.push(p);
      p.styleBuiltIn = Word.BuiltInStyleName.normal;
      p.load('isListItem');
      await context.sync();
      if (p.isListItem) {
        p.detachFromList();
        await context.sync();
        notes.push('the new paragraph at the end of the document inherited list membership (detached first)');
      }
      return p;
    }
  };

  let check: SelfTestCheck;
  try {
    const outcome = await body(scope);
    await context.sync();
    check = { label, pass: outcome.pass, detail: [outcome.detail, ...notes].join('; '), info: options.info };
  } catch (err) {
    check = { label, pass: false, detail: [`threw: ${describeError(err)}`, ...notes].join('; '), info: options.info };
  }

  // Delete what this probe created, newest first, one sync each so one
  // failure (e.g. an object that was never created) doesn't block the rest.
  let cleanupFailures = 0;
  for (const obj of created.reverse()) {
    try {
      obj.delete();
      await context.sync();
    } catch {
      cleanupFailures++;
    }
  }
  if (cleanupFailures) check.detail = `${check.detail}; cleanup: ${cleanupFailures} object(s) could not be deleted`;
  return check;
}

async function paragraphCount(context: Word.RequestContext): Promise<number> {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load('items/text');
  await context.sync();
  return paragraphs.items.length;
}

async function runCapabilityProbes(context: Word.RequestContext, onProbe: (label: string) => void): Promise<SelfTestCheck[]> {
  const checks: SelfTestCheck[] = [];
  const probe = async (label: string, options: { info?: boolean }, body: (s: ProbeScope) => Promise<ProbeOutcome>): Promise<void> => {
    onProbe(label);
    checks.push(await runProbe(context, label, options, body));
  };

  let countBefore: number | null = null;
  try {
    countBefore = await paragraphCount(context);
  } catch {
    countBefore = null;
  }

  await probe('insert paragraph + styleBuiltIn Heading2', {}, async (s) => {
    const p = await s.plainParagraph('Probe: heading');
    p.styleBuiltIn = Word.BuiltInStyleName.heading2;
    p.load('styleBuiltIn');
    await context.sync();
    return { pass: p.styleBuiltIn === 'Heading2', detail: `styleBuiltIn=${p.styleBuiltIn}` };
  });

  await probe('styleBuiltIn ListParagraph on a plain paragraph (no list)', {}, async (s) => {
    const p = await s.plainParagraph('Probe: list paragraph style');
    p.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: f.styleBuiltIn === 'ListParagraph' && !f.isListItem, detail: describeFacts(f) };
  });

  await probe('startNewList() on a paragraph with text', {}, async (s) => {
    const p = await s.plainParagraph('Probe: list start with text');
    p.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    p.startNewList();
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: f.isListItem, detail: `${describeFacts(f)} (Word's default list indent)` };
  });

  await probe('startNewList() on an empty paragraph', {}, async (s) => {
    const p = await s.plainParagraph('');
    p.startNewList();
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: f.isListItem, detail: describeFacts(f) };
  });

  await probe('list.setLevelBullet(0, solid)', {}, async (s) => {
    const p = await s.plainParagraph('Probe: bullet level');
    const list = p.startNewList();
    await context.sync();
    list.setLevelBullet(0, Word.ListBullet.solid);
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: f.isListItem, detail: `listString=${JSON.stringify(f.listString)}` };
  });

  await probe('list.setLevelNumbering(0, arabic, [0, "."]) shows "1."', {}, async (s) => {
    const p = await s.plainParagraph('Probe: number level');
    const list = p.startNewList();
    await context.sync();
    list.setLevelNumbering(0, Word.ListNumbering.arabic, [0, '.']);
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: (f.listString ?? '').trim() === '1.', detail: `listString=${JSON.stringify(f.listString)}` };
  });

  await probe('listItem.insertParagraph("x", "After"): does the new paragraph inherit list membership?', { info: true }, async (s) => {
    const p = await s.plainParagraph('Probe: list item before');
    const list = p.startNewList();
    list.load('id');
    await context.sync();
    const x = p.insertParagraph('Probe: inserted after a list item', Word.InsertLocation.after);
    s.created(x);
    await context.sync();
    const f = await readFacts(context, x);
    return {
      pass: true,
      detail: `isListItem=${f.isListItem}${f.isListItem ? `, same list as the item before: ${f.listId === list.id}` : ''}, style=${f.styleBuiltIn}`
    };
  });

  await probe('attachToList(list.id, 0) on a plain paragraph with text → list item of that list', {}, async (s) => {
    const owner = await s.plainParagraph('Probe: list owner');
    owner.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    const list = owner.startNewList();
    list.load('id');
    await context.sync();
    const p = await s.plainParagraph('Probe: attach at level 0');
    p.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    p.attachToList(list.id, 0);
    await context.sync();
    const f = await readFacts(context, p);
    return {
      pass: f.isListItem && f.listId === list.id,
      detail: `list.id=${list.id}; ${describeFacts(f)}`
    };
  });

  await probe('attachToList(list.id, 1) → listItem.level == 1', {}, async (s) => {
    const owner = await s.plainParagraph('Probe: list owner');
    owner.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    const list = owner.startNewList();
    list.load('id');
    await context.sync();
    const p = await s.plainParagraph('Probe: attach at level 1');
    p.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    p.attachToList(list.id, 1);
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: f.isListItem && f.level === 1, detail: `list.id=${list.id}; ${describeFacts(f)}` };
  });

  await probe('detachFromList() on a list item → not a list item', {}, async (s) => {
    const p = await s.plainParagraph('Probe: detach');
    p.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
    p.startNewList();
    await context.sync();
    p.detachFromList();
    await context.sync();
    const f = await readFacts(context, p);
    return { pass: !f.isListItem, detail: describeFacts(f) };
  });

  const bullet0 = listLevelIndents(0);
  for (const bulletArg of [bullet0.bulletIndent, bullet0.bulletIndent - bullet0.textIndent]) {
    const asUsed = bulletArg === (bullet0.bulletIndent - bullet0.textIndent);
    await probe(
      `list.setLevelIndents(0, ${bullet0.textIndent}, ${bulletArg})${asUsed ? ' (as used by insert)' : ''}`,
      { info: !asUsed },
      async (s) => {
        const p = await s.plainParagraph('Probe: level indents');
        p.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
        const list = p.startNewList();
        await context.sync();
        list.setLevelBullet(0, Word.ListBullet.solid);
        await context.sync();
        list.setLevelIndents(0, bullet0.textIndent, bulletArg);
        await context.sync();
        const f = await readFacts(context, p);
        const matches =
          Math.abs(f.leftIndent - bullet0.textIndent) <= 1.5 && Math.abs(f.leftIndent + f.firstLineIndent - bullet0.bulletIndent) <= 1.5;
        return {
          pass: true,
          detail: `${indentFacts(f)} → ${matches ? "matches Word's bullet button" : `does NOT match Word's bullet button (${bullet0.textIndent}/${bullet0.bulletIndent})`}`
        };
      }
    );
  }

  for (const setStyleAfter of [false, true]) {
    await probe(
      `insertOoxml bullet item, then getOoxml(): <w:numPr>/<w:pStyle> present? (${setStyleAfter ? 'styleBuiltIn set afterwards' : 'no style set afterwards'})`,
      { info: true },
      async (s) => {
        const holder = await s.plainParagraph('');
        const xml = buildListOoxml([{ level: 0, ordered: false, inlines: [{ text: 'Probe: ooxml item' }] }]);
        const range = holder.insertOoxml(xml, Word.InsertLocation.replace);
        const paragraphs = range.paragraphs;
        paragraphs.load('items/text');
        await context.sync();
        paragraphs.items.forEach((p) => s.created(p));
        const first = paragraphs.items[0];
        if (!first) return { pass: false, detail: 'insertOoxml produced no paragraph' };
        if (setStyleAfter) {
          first.styleBuiltIn = Word.BuiltInStyleName.listParagraph;
          await context.sync();
        }
        const ooxml = first.getOoxml();
        await context.sync();
        const facts = ooxmlBodyFacts(ooxml.value);
        const f = await readFacts(context, first);
        return {
          pass: true,
          detail:
            `${paragraphs.items.length} paragraph(s); isListItem=${f.isListItem}, style=${f.styleBuiltIn}; ` +
            `<w:numPr> in body: ${facts.numPr ? 'yes' : 'no'}; <w:pStyle>: ${facts.pStyle === null ? 'no' : `yes (w:val="${facts.pStyle}")`}`
        };
      }
    );
  }

  await probe('insertTable 2x2 + styleBuiltIn GridTable4 + headerRowCount 1', {}, async (s) => {
    const holder = await s.plainParagraph('');
    const table = holder.insertTable(2, 2, Word.InsertLocation.after, [
      ['A', 'B'],
      ['1', '2']
    ]);
    s.created(table);
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable4;
    table.headerRowCount = 1;
    table.load('styleBuiltIn,headerRowCount,rowCount');
    await context.sync();
    return {
      pass: table.headerRowCount === 1,
      detail: `styleBuiltIn=${table.styleBuiltIn}, headerRowCount=${table.headerRowCount}, rowCount=${table.rowCount}`
    };
  });

  await probe('table look: GridTable4 + house flags → tblLook firstColumn=1, lastColumn=0?', {}, async (s) => {
    const holder = await s.plainParagraph('');
    const table = holder.insertTable(3, 2, Word.InsertLocation.after, [
      ['Col A', 'Col B'],
      ['a1', 'b1'],
      ['a2', 'b2']
    ]);
    s.created(table);
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable4;
    table.headerRowCount = 1;
    table.styleFirstColumn = true;
    table.styleBandedRows = true;
    table.styleBandedColumns = false;
    table.styleLastColumn = false;
    table.styleTotalRow = false;
    await context.sync();

    // Read back the API values
    table.load('styleFirstColumn,styleLastColumn,styleBandedRows,styleBandedColumns,styleTotalRow,headerRowCount');
    await context.sync();

    // Read the OOXML to get the actual tblLook
    const range = table.getRange();
    const ooxml = range.getOoxml();
    await context.sync();

    // Extract tblLook from OOXML using regex
    const tblLookMatch = /<w:tblLook[^>]*w:val="([0-9A-Fa-f]+)"/.exec(ooxml.value);
    const hexVal = tblLookMatch ? parseInt(tblLookMatch[1], 16) : 0;
    const firstColumn = !!(hexVal & 0x0080);
    const lastColumn = !!(hexVal & 0x0100);

    // Read bold status of cells (0,0), (0,1), (1,0), (1,1)
    const boldStatus: { [key: string]: boolean | string } = {};
    for (let r = 0; r < 2; r++) {
      for (let c = 0; c < 2; c++) {
        try {
          const cell = table.getCell(r, c);
          const range = cell.body.getRange();
          range.load('font/bold');
          await context.sync();
          boldStatus[`bold(${r},${c})`] = range.font.bold;
        } catch {
          boldStatus[`bold(${r},${c})`] = 'error';
        }
      }
    }

    const pass = firstColumn === true && lastColumn === false;
    const detail = [
      `API: styleFirstColumn=${table.styleFirstColumn}, styleLastColumn=${table.styleLastColumn}, styleBandedRows=${table.styleBandedRows}, styleBandedColumns=${table.styleBandedColumns}`,
      `tblLook hex=${tblLookMatch ? tblLookMatch[1] : 'not found'} → firstColumn=${firstColumn}, lastColumn=${lastColumn}`,
      `cell formatting: ${Object.entries(boldStatus)
        .map(([k, v]) => `${k}=${v}`)
        .join(', ')}`
    ].join('; ');

    return { pass, detail };
  });

  await probe('selection.isEmpty can be loaded (a collapsed cursor is never deleted)', {}, async () => {
    const selection = context.document.getSelection();
    selection.load('isEmpty');
    await context.sync();
    // Mirrors getCursor: only a non-empty selection is ever deleted, and
    // this probe leaves Kevin's selection alone either way.
    return { pass: true, detail: `isEmpty=${selection.isEmpty}` };
  });

  onProbe('cleanup check');
  try {
    const countAfter = await paragraphCount(context);
    let tidied = 0;
    if (countBefore !== null && countAfter > countBefore) {
      // Probes only ever append at the end; remove empty leftovers beyond the original count.
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load('items/text');
      await context.sync();
      for (const p of paragraphs.items.slice(countBefore).reverse()) {
        if (p.text.trim() !== '') continue;
        try {
          p.delete();
          await context.sync();
          tidied++;
        } catch {
          /* the document's final paragraph mark can't be deleted */
        }
      }
    }
    const countFinal = tidied ? await paragraphCount(context) : countAfter;
    const unchanged = countBefore !== null && countFinal === countBefore;
    checks.push({
      label: 'probe cleanup left the document as it was',
      pass: true,
      info: true,
      detail:
        `${unchanged ? 'unchanged' : 'CHANGED'}: paragraphs before: ${countBefore ?? '?'}, after probes: ${countAfter}` +
        (tidied ? `, after removing ${tidied} empty leftover(s): ${countFinal}` : '')
    });
  } catch (err) {
    checks.push({ label: 'probe cleanup left the document as it was', pass: false, info: true, detail: `threw: ${describeError(err)}` });
  }

  return checks;
}
