import type { Block } from '../model';
import { parseMarkdown } from '../parse/markdown';
import { parseHtml } from '../parse/html';
import { detectFormat, formatLabel, type DetectedFormat } from '../parse/detect';
import { normalize, type NormalizeOptions } from '../transform/normalize';
import { DEFAULT_TABLE_STYLE, InsertError, type InsertOptions, type StepLogEntry } from '../word/insert';
import {
  loadNormalizeOptions,
  saveNormalizeOptions,
  loadStyleChoice,
  saveStyleChoice
} from './optionsStore';

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

interface State {
  rawText: string;
  rawHtml?: string;
  formatOverride: DetectedFormat | null;
  normalizeOptions: NormalizeOptions;
  bodyStyleChoice: string;
  tableStyleChoice: string;
  blocks: Block[];
  officeReady: boolean;
  wordApiOk: boolean;
}

const state: State = {
  rawText: '',
  rawHtml: undefined,
  formatOverride: null,
  normalizeOptions: loadNormalizeOptions(),
  bodyStyleChoice: loadStyleChoice('body'),
  tableStyleChoice: loadStyleChoice('table'),
  blocks: [],
  officeReady: false,
  wordApiOk: false
};

// ---------------------------------------------------------------------------
// DOM handles
// ---------------------------------------------------------------------------

const el = {
  hostBanner: byId('host-banner'),
  pasteBox: byId('paste-box') as HTMLTextAreaElement,
  pasteNotice: byId('paste-notice'),
  detectedFormat: byId('detected-format'),
  formatOverride: byId('format-override'),
  preview: byId('preview'),
  optTopHeading: byId('opt-top-heading') as HTMLSelectElement,
  optBoldHeading: byId('opt-bold-heading') as HTMLInputElement,
  optCitations: byId('opt-citations') as HTMLInputElement,
  optManualNumbers: byId('opt-manual-numbers') as HTMLInputElement,
  optEmoji: byId('opt-emoji') as HTMLInputElement,
  optBlankLines: byId('opt-blank-lines') as HTMLInputElement,
  optBodyStyle: byId('opt-body-style') as HTMLSelectElement,
  optTableStyle: byId('opt-table-style') as HTMLSelectElement,
  insertBtn: byId('insert-btn') as HTMLButtonElement,
  clearBtn: byId('clear-btn') as HTMLButtonElement,
  status: byId('status'),
  errorDetails: byId('error-details'),
  errorDetailsText: byId('error-details-text'),
  copyErrorBtn: byId('copy-error-btn') as HTMLButtonElement,
  copyErrorFeedback: byId('copy-error-feedback'),
  selftestLink: byId('selftest-link'),
  selftestResult: byId('selftest-result'),
  buildVersion: byId('build-version')
};

/** The full diagnostic payload behind the current "Copy details" button, if any. */
let lastCopyPayload: string | null = null;

function byId(id: string): HTMLElement {
  const found = document.getElementById(id);
  if (!found) throw new Error(`Missing #${id} in taskpane.html`);
  return found;
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------

init();

async function init(): Promise<void> {
  el.buildVersion.textContent = `AI Paste v${__APP_VERSION__} (${__GIT_HASH__})`;
  hydrateOptionControls();
  wireEvents();

  if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
    try {
      await Office.onReady();
      state.officeReady = true;
    } catch {
      state.officeReady = false;
    }
  }

  if (state.officeReady) {
    const { isWordApi13Supported, loadDocumentStyleNames } = await import('../word/officeAdapter');
    state.wordApiOk = isWordApi13Supported();
    if (!state.wordApiOk) {
      showHostBanner(
        'AI Paste needs to run inside Word for Microsoft 365 (or another host with WordApi 1.3+). Pasting and preview still work here, but "Insert at cursor" is disabled.'
      );
    } else {
      try {
        const { paragraphStyles, tableStyles } = await loadDocumentStyleNames();
        populateStyleOptions(el.optBodyStyle, paragraphStyles, 'custom');
        populateStyleOptions(el.optTableStyle, tableStyles, 'custom');
        el.optBodyStyle.value = state.bodyStyleChoice;
        el.optTableStyle.value = state.tableStyleChoice;
      } catch {
        /* best-effort; built-in options still work */
      }
    }
  } else {
    showHostBanner('Not running inside Word — you can still paste and preview here. Open this pane from Word to insert.');
  }

  updateInsertAvailability();
  runPipeline();
}

function showHostBanner(message: string): void {
  el.hostBanner.textContent = message;
  el.hostBanner.hidden = false;
}

function populateStyleOptions(select: HTMLSelectElement, names: string[], prefix: 'custom'): void {
  for (const name of names) {
    const opt = document.createElement('option');
    opt.value = `${prefix}:${name}`;
    opt.textContent = name;
    select.appendChild(opt);
  }
}

function hydrateOptionControls(): void {
  el.optTopHeading.value = String(state.normalizeOptions.topHeadingLevel);
  el.optBoldHeading.checked = state.normalizeOptions.boldLineToHeading;
  el.optCitations.checked = state.normalizeOptions.removeCitations;
  el.optManualNumbers.checked = state.normalizeOptions.removeManualHeadingNumbers;
  el.optEmoji.checked = state.normalizeOptions.removeEmoji;
  el.optBlankLines.checked = state.normalizeOptions.blankLinesBetweenBlocks;
  el.optBodyStyle.value = state.bodyStyleChoice;
  el.optTableStyle.value = state.tableStyleChoice;
}

function updateInsertAvailability(): void {
  el.insertBtn.disabled = !state.officeReady || !state.wordApiOk || state.blocks.length === 0;
}

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

function wireEvents(): void {
  el.pasteBox.addEventListener('paste', onPaste);
  el.pasteBox.addEventListener('input', onManualEdit);

  el.formatOverride.querySelectorAll<HTMLButtonElement>('[data-format]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.formatOverride = btn.dataset.format as DetectedFormat;
      runPipeline();
    });
  });

  el.optTopHeading.addEventListener('change', onNormalizeOptionChange);
  el.optBoldHeading.addEventListener('change', onNormalizeOptionChange);
  el.optCitations.addEventListener('change', onNormalizeOptionChange);
  el.optManualNumbers.addEventListener('change', onNormalizeOptionChange);
  el.optEmoji.addEventListener('change', onNormalizeOptionChange);
  el.optBlankLines.addEventListener('change', onNormalizeOptionChange);

  el.optBodyStyle.addEventListener('change', () => {
    state.bodyStyleChoice = el.optBodyStyle.value;
    saveStyleChoice('body', state.bodyStyleChoice);
  });
  el.optTableStyle.addEventListener('change', () => {
    state.tableStyleChoice = el.optTableStyle.value;
    saveStyleChoice('table', state.tableStyleChoice);
  });

  el.insertBtn.addEventListener('click', onInsertClick);
  el.clearBtn.addEventListener('click', onClearClick);
  el.selftestLink.addEventListener('click', onSelfTestClick);
  el.copyErrorBtn.addEventListener('click', onCopyDetailsClick);
}

async function onCopyDetailsClick(): Promise<void> {
  if (!lastCopyPayload) return;
  try {
    await navigator.clipboard.writeText(lastCopyPayload);
    el.copyErrorFeedback.hidden = false;
    setTimeout(() => {
      el.copyErrorFeedback.hidden = true;
    }, 2000);
  } catch {
    // Clipboard API can be blocked in some hosts; the pre-formatted text is
    // still selectable/visible on screen either way.
  }
}

function onPaste(e: ClipboardEvent): void {
  const cd = e.clipboardData;
  if (!cd) return;
  const html = cd.getData('text/html');
  const text = cd.getData('text/plain');
  if (!html && !text) return;

  e.preventDefault();
  state.rawText = text;
  state.rawHtml = html || undefined;
  state.formatOverride = null;
  el.pasteBox.value = text;
  el.pasteNotice.hidden = true;
  runPipeline();
}

function onManualEdit(): void {
  const current = el.pasteBox.value;
  if (state.rawHtml && current !== state.rawText) {
    state.rawHtml = undefined;
    el.pasteNotice.textContent = 'Switched to plain-text mode after manual edits.';
    el.pasteNotice.hidden = false;
  }
  state.rawText = current;
  runPipeline();
}

function onNormalizeOptionChange(): void {
  state.normalizeOptions = {
    topHeadingLevel: Number(el.optTopHeading.value) || 1,
    boldLineToHeading: el.optBoldHeading.checked,
    removeCitations: el.optCitations.checked,
    removeManualHeadingNumbers: el.optManualNumbers.checked,
    removeEmoji: el.optEmoji.checked,
    blankLinesBetweenBlocks: el.optBlankLines.checked
  };
  saveNormalizeOptions(state.normalizeOptions);
  runPipeline();
}

function onClearClick(): void {
  state.rawText = '';
  state.rawHtml = undefined;
  state.formatOverride = null;
  el.pasteBox.value = '';
  el.pasteNotice.hidden = true;
  setStatus('', 'idle');
  hideErrorDetails();
  runPipeline();
}

async function onInsertClick(): Promise<void> {
  if (!state.blocks.length) return;
  el.insertBtn.disabled = true;
  hideErrorDetails();
  setStatus('Inserting…', 'idle');
  try {
    const { insertBlocksInWord } = await import('../word/officeAdapter');
    await insertBlocksInWord(state.blocks, buildInsertOptions());
    setStatus(`Inserted ${state.blocks.length} block${state.blocks.length === 1 ? '' : 's'}.`, 'ok');
  } catch (err) {
    if (err instanceof InsertError) {
      setStatus(`Insert failed: ${err.officeError.message}`, 'error');
      showErrorDetails(err);
    } else {
      setStatus(`Insert failed: ${errorMessage(err)}`, 'error');
    }
  } finally {
    updateInsertAvailability();
  }
}

function hideErrorDetails(): void {
  el.errorDetails.hidden = true;
  el.errorDetailsText.textContent = '';
  lastCopyPayload = null;
}

function showErrorDetails(err: InsertError): void {
  const step = err.failedStep;
  const lines = [
    step ? `Failed step: block ${step.blockIndex} (${step.blockType}): ${step.action}` : 'Failed during setup',
    `code: ${err.officeError.code ?? '(none)'}`,
    `message: ${err.officeError.message}`,
    `errorLocation: ${err.officeError.errorLocation ?? '(none)'}`,
    `statement: ${err.officeError.statement ?? '(none)'}`
  ];
  el.errorDetailsText.textContent = lines.join('\n');
  el.errorDetails.hidden = false;
  el.copyErrorFeedback.hidden = true;
  lastCopyPayload = formatStepLogForCopy(err.officeError, step, err.steps);
}

function formatStepLogForCopy(
  officeError: { code?: string; message: string; errorLocation?: string; statement?: string },
  failedStep: StepLogEntry | null,
  steps: StepLogEntry[]
): string {
  const payload = {
    version: `${__APP_VERSION__} (${__GIT_HASH__})`,
    when: new Date().toISOString(),
    officeError,
    failedStep,
    steps: steps.map((s) => `block ${s.blockIndex} ${s.blockType}: ${s.action}`)
  };
  return JSON.stringify(payload, null, 2);
}

interface SelfTestCheckLike {
  label: string;
  pass: boolean;
  detail?: string;
}
interface SelfTestResultLike {
  ranAt: string;
  allPassed: boolean;
  probes?: SelfTestCheckLike[];
  checks: SelfTestCheckLike[];
}

async function onSelfTestClick(e: Event): Promise<void> {
  e.preventDefault();
  if (!state.officeReady || !state.wordApiOk) {
    renderSelfTestResult({
      ranAt: new Date().toISOString(),
      allPassed: false,
      checks: [{ label: 'Running inside Word with WordApi 1.3', pass: false, detail: 'Open this pane inside Word to run the self-test.' }]
    });
    return;
  }
  el.selftestLink.textContent = 'Running self-test…';
  try {
    const { runSelfTest } = await import('../word/selftest');
    const result = await runSelfTest();
    renderSelfTestResult(result);
  } catch (err) {
    renderSelfTestResult({
      ranAt: new Date().toISOString(),
      allPassed: false,
      checks: [{ label: 'Self-test ran without throwing', pass: false, detail: errorMessage(err) }]
    });
  } finally {
    el.selftestLink.textContent = 'Run self-test';
  }
}

function renderSelfTestResult(result: SelfTestResultLike): void {
  el.selftestResult.hidden = false;

  const list = (checks: SelfTestCheckLike[]): string =>
    `<ul>${checks
      .map(
        (c) =>
          `<li class="${c.pass ? 'selftest-pass' : 'selftest-fail'}">${c.pass ? '✓' : '✗'} ${escapeHtml(c.label)}${
            c.detail ? ` — <span class="muted">${escapeHtml(c.detail)}</span>` : ''
          }</li>`
      )
      .join('')}</ul>`;

  const probesHtml = result.probes?.length ? `<strong>Capability probes</strong>${list(result.probes)}` : '';
  const checksHtml = `<strong>Fixture</strong>${list(result.checks)}`;

  el.selftestResult.innerHTML =
    `<strong>${result.allPassed ? 'All checks passed' : 'Some checks failed'}</strong>` +
    probesHtml +
    checksHtml +
    `<button id="selftest-copy-btn" class="btn btn-secondary btn-small" type="button">Copy results</button>` +
    `<span id="selftest-copy-feedback" class="copy-feedback" hidden>Copied.</span>`;

  const copyBtn = document.getElementById('selftest-copy-btn') as HTMLButtonElement | null;
  const feedback = document.getElementById('selftest-copy-feedback');
  copyBtn?.addEventListener('click', async () => {
    const payload = JSON.stringify({ version: `${__APP_VERSION__} (${__GIT_HASH__})`, ...result }, null, 2);
    try {
      await navigator.clipboard.writeText(payload);
      if (feedback) {
        feedback.hidden = false;
        setTimeout(() => {
          feedback.hidden = true;
        }, 2000);
      }
    } catch {
      /* clipboard API can be blocked; the checklist is still visible on screen */
    }
  });
}

// ---------------------------------------------------------------------------
// Pipeline: paste -> detect -> parse -> normalize -> preview
// ---------------------------------------------------------------------------

function runPipeline(): void {
  const effectiveFormat = state.formatOverride ?? detectFormat({ text: state.rawText, html: state.rawHtml });
  el.detectedFormat.textContent = formatLabel(effectiveFormat);
  el.formatOverride.querySelectorAll<HTMLButtonElement>('[data-format]').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.format === effectiveFormat);
  });

  let parsed: Block[] = [];
  try {
    if (effectiveFormat === 'html' && state.rawHtml) {
      parsed = parseHtml(state.rawHtml);
    } else if (state.rawText.trim()) {
      parsed = parseMarkdown(state.rawText);
    }
  } catch {
    parsed = [];
  }

  state.blocks = normalize(parsed, state.normalizeOptions);
  renderPreview(state.blocks);
  updateInsertAvailability();
}

function renderPreview(blocks: Block[]): void {
  if (!blocks.length) {
    el.preview.innerHTML = '<p class="muted">Nothing pasted yet.</p>';
    return;
  }

  const headingLines = blocks
    .filter((b) => b.type === 'heading')
    .map((b) => {
      const indent = ' '.repeat(Math.max(0, b.level - 1));
      const text = b.inlines.map((i) => i.text).join('') || '(untitled)';
      return `<li>${indent}${escapeHtml(text)}</li>`;
    })
    .join('');

  const paragraphCount = blocks.filter((b) => b.type === 'paragraph' && b.inlines.length > 0).length;
  const listCount = new Set(blocks.filter((b) => b.type === 'listItem').map((b) => b.listIndex)).size;
  const tableCount = blocks.filter((b) => b.type === 'table').length;

  const counts = `${paragraphCount} paragraph${paragraphCount === 1 ? '' : 's'} · ${listCount} list${listCount === 1 ? '' : 's'} · ${tableCount} table${tableCount === 1 ? '' : 's'}`;

  el.preview.innerHTML =
    (headingLines ? `<ul class="preview-outline">${headingLines}</ul>` : '') + `<p class="preview-counts">${counts}</p>`;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function buildInsertOptions(): InsertOptions {
  return {
    bodyStyle: decodeChoice(state.bodyStyleChoice) === 'builtin' ? { kind: 'builtin' } : { kind: 'custom', name: choiceName(state.bodyStyleChoice) },
    tableStyle: decodeTableStyle(state.tableStyleChoice)
  };
}

function decodeChoice(value: string): 'builtin' | 'custom' {
  return value.startsWith('custom:') ? 'custom' : 'builtin';
}

function choiceName(value: string): string {
  return value.slice(value.indexOf(':') + 1);
}

function decodeTableStyle(value: string) {
  if (decodeChoice(value) === 'custom') {
    return { ...DEFAULT_TABLE_STYLE, styleBuiltIn: null, customName: choiceName(value) };
  }
  const name = choiceName(value);
  if (name === 'GridTable1Light') {
    return { ...DEFAULT_TABLE_STYLE, styleBuiltIn: 'GridTable1Light', styleFirstColumn: false, styleBandedRows: false };
  }
  if (name === 'PlainTable1') {
    return { ...DEFAULT_TABLE_STYLE, styleBuiltIn: 'PlainTable1', styleFirstColumn: false, styleBandedRows: false };
  }
  return { ...DEFAULT_TABLE_STYLE, styleBuiltIn: 'GridTable4' };
}

function setStatus(message: string, kind: 'idle' | 'ok' | 'error'): void {
  el.status.textContent = message;
  el.status.className = `status${kind === 'ok' ? ' status-ok' : kind === 'error' ? ' status-error' : ''}`;
}

function errorMessage(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (typeof err === 'object' && err && 'message' in err) return String((err as { message: unknown }).message);
  return String(err);
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c] as string);
}
