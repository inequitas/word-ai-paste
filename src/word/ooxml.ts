import type { Inline } from '../model';

/**
 * Builds a minimal flat-OPC WordprocessingML package for `insertOoxml`, used
 * as a fallback for inserting a whole list at once when the paragraph/list
 * API path (`startNewList`, `list.insertParagraph`, `setLevelBullet`, …)
 * throws on a given Word host. Pure string building — no Office.js — so it
 * can be unit-tested directly (see tests/ooxml.test.ts) by parsing the
 * result as XML.
 *
 * Deliberately minimal: no fonts, sizes, or colors — only `<w:b/>`/`<w:i/>`
 * on runs that need them, `pStyle="ListParagraph"` (matching the document's
 * own "List Paragraph" style by its portable name, the same one
 * `styleBuiltIn` uses elsewhere), and one plain numbering definition. Each
 * insert uses a fresh `numId`, so Word starts a new list instance and
 * numbering restarts at 1, same as the primary API path.
 */

export interface OoxmlListItem {
  /** 0-based nesting level. */
  level: number;
  ordered: boolean;
  inlines: Inline[];
}

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG_NS = 'http://schemas.microsoft.com/office/2006/xmlPackage';
const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

// Word's own default bullet set, repeating every 3 levels — the same glyphs
// Word's bullet button uses (see the sample docs' numbering.xml, read
// locally during development: level 0 Symbol U+F0B7, level 1 "o" in
// Courier New, level 2 Wingdings U+F0A7).
const BULLET_GLYPHS: { char: string; font: string }[] = [
  { char: '', font: 'Symbol' },
  { char: 'o', font: 'Courier New' },
  { char: '', font: 'Wingdings' }
];

export function buildListOoxml(items: OoxmlListItem[]): string {
  const bodyXml = items.map(itemParagraphXml).join('');
  const documentXml = `<w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document>`;
  return wrapFlatOpc({
    documentXml,
    numberingXml: buildNumberingXml(items),
    stylesXml: buildStylesXml()
  });
}

function itemParagraphXml(item: OoxmlListItem): string {
  const runsXml = item.inlines.map(runXml).join('') || `<w:r/>`;
  return (
    `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>` +
    `<w:numPr><w:ilvl w:val="${item.level}"/><w:numId w:val="1"/></w:numPr>` +
    `</w:pPr>${runsXml}</w:p>`
  );
}

function runXml(inline: Inline): string {
  if (!inline.text) return '';
  const rPr = rprXml(inline);
  const segments = inline.text.split('\n');
  const inner = segments
    .map((seg, i) => (i > 0 ? '<w:br/>' : '') + (seg ? `<w:t xml:space="preserve">${escapeXml(seg)}</w:t>` : ''))
    .join('');
  return `<w:r>${rPr}${inner}</w:r>`;
}

function rprXml(inline: Inline): string {
  const bold = inline.bold ? '<w:b/>' : '';
  const italic = inline.italic ? '<w:i/>' : '';
  if (!bold && !italic) return '';
  return `<w:rPr>${bold}${italic}</w:rPr>`;
}

function buildNumberingXml(items: OoxmlListItem[]): string {
  const levelOrdered = new Map<number, boolean>();
  for (const item of items) {
    if (!levelOrdered.has(item.level)) levelOrdered.set(item.level, item.ordered);
  }
  const defaultOrdered = levelOrdered.get(0) ?? false;

  const levels: string[] = [];
  for (let level = 0; level < 9; level++) {
    levels.push(levelXml(level, levelOrdered.get(level) ?? defaultOrdered));
  }

  return (
    `<w:numbering xmlns:w="${W_NS}">` +
    `<w:abstractNum w:abstractNumId="1">${levels.join('')}</w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>` +
    `</w:numbering>`
  );
}

function levelXml(level: number, ordered: boolean): string {
  const indentLeft = 720 + 720 * level;
  const hanging = 360;
  const indentXml = `<w:ind w:left="${indentLeft}" w:hanging="${hanging}"/>`;

  if (ordered) {
    return (
      `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
      `<w:lvlText w:val="%${level + 1}."/><w:lvlJc w:val="left"/>` +
      `<w:pPr>${indentXml}</w:pPr></w:lvl>`
    );
  }

  const glyph = BULLET_GLYPHS[level % BULLET_GLYPHS.length];
  return (
    `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
    `<w:lvlText w:val="${escapeXml(glyph.char)}"/><w:lvlJc w:val="left"/>` +
    `<w:pPr>${indentXml}</w:pPr>` +
    `<w:rPr><w:rFonts w:ascii="${glyph.font}" w:hAnsi="${glyph.font}" w:hint="default"/></w:rPr></w:lvl>`
  );
}

function buildStylesXml(): string {
  return (
    `<w:styles xmlns:w="${W_NS}">` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="ListParagraph">` +
    `<w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>` +
    `<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>` +
    `</w:style>` +
    `</w:styles>`
  );
}

function wrapFlatOpc(parts: { documentXml: string; numberingXml: string; stylesXml: string }): string {
  const relsRoot =
    `<Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;
  const relsDoc =
    `<Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`;

  const part = (name: string, contentType: string, xml: string): string =>
    `<pkg:part pkg:name="${name}" pkg:contentType="${contentType}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pkg:package xmlns:pkg="${PKG_NS}">` +
    part('/_rels/.rels', 'application/vnd.openxmlformats-package.relationships+xml', relsRoot) +
    part(
      '/word/document.xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
      parts.documentXml
    ) +
    part('/word/_rels/document.xml.rels', 'application/vnd.openxmlformats-package.relationships+xml', relsDoc) +
    part(
      '/word/numbering.xml',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
      parts.numberingXml
    ) +
    part('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml', parts.stylesXml) +
    `</pkg:package>`
  );
}

export function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}
