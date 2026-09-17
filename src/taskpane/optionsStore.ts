import { DEFAULT_NORMALIZE_OPTIONS, type NormalizeOptions } from '../transform/normalize';

const NORMALIZE_KEY = 'ai-paste:normalize-options:v1';
const BODY_STYLE_KEY = 'ai-paste:body-style:v1';
const TABLE_STYLE_KEY = 'ai-paste:table-style:v1';

export const DEFAULT_BODY_STYLE_CHOICE = 'builtin:Normal';
export const DEFAULT_TABLE_STYLE_CHOICE = 'builtin:GridTable4';

/** All persistence goes through localStorage, wrapped in try/catch — a private window or blocked
 * site data must never break the add-in, it should just fall back to defaults every time. */

export function loadNormalizeOptions(): NormalizeOptions {
  try {
    const raw = localStorage.getItem(NORMALIZE_KEY);
    if (!raw) return { ...DEFAULT_NORMALIZE_OPTIONS };
    const parsed = JSON.parse(raw) as Partial<NormalizeOptions>;
    return { ...DEFAULT_NORMALIZE_OPTIONS, ...parsed };
  } catch {
    return { ...DEFAULT_NORMALIZE_OPTIONS };
  }
}

export function saveNormalizeOptions(options: NormalizeOptions): void {
  try {
    localStorage.setItem(NORMALIZE_KEY, JSON.stringify(options));
  } catch {
    /* ignore */
  }
}

export function loadStyleChoice(kind: 'body' | 'table'): string {
  const key = kind === 'body' ? BODY_STYLE_KEY : TABLE_STYLE_KEY;
  const fallback = kind === 'body' ? DEFAULT_BODY_STYLE_CHOICE : DEFAULT_TABLE_STYLE_CHOICE;
  try {
    return localStorage.getItem(key) || fallback;
  } catch {
    return fallback;
  }
}

export function saveStyleChoice(kind: 'body' | 'table', value: string): void {
  const key = kind === 'body' ? BODY_STYLE_KEY : TABLE_STYLE_KEY;
  try {
    localStorage.setItem(key, value);
  } catch {
    /* ignore */
  }
}
