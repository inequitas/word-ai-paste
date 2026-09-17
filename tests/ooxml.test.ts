import { describe, it, expect } from 'vitest';
import { DOMParser as XmlDOMParser } from '@xmldom/xmldom';
import { buildListOoxml, escapeXml, type OoxmlListItem } from '../src/word/ooxml';

// happy-dom's own DOMParser doesn't implement real XML parsing (it always
// parses as HTML regardless of the requested mimeType — confirmed by
// hand — which would silently lowercase/mangle namespaced tags like
// <w:p>). @xmldom/xmldom is a real, dependency-free XML DOMParser, used
// here only to prove the OOXML builder's output is well-formed XML with
// correct namespaces.
let parserErrors: string[] = [];
function parseXml(xml: string): Document {
  parserErrors = [];
  const doc = new XmlDOMParser({
    onError: (_level: string, msg: string) => parserErrors.push(msg)
  }).parseFromString(xml, 'application/xml');
  if (parserErrors.length) throw new Error('XML did not parse: ' + parserErrors.join('; '));
  return doc as unknown as Document;
}

const NS = {
  pkg: 'http://schemas.microsoft.com/office/2006/xmlPackage',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
};

function firstElement(node: Element): Element {
  // @xmldom/xmldom doesn't implement the DOM Level 4 `firstElementChild`
  // convenience property, only classic `childNodes`.
  const found = Array.from(node.childNodes).find((n) => n.nodeType === 1) as Element | undefined;
  if (!found) throw new Error('no element child found');
  return found;
}

function partXmlData(doc: Document, partName: string): Element {
  const part = Array.from(doc.getElementsByTagNameNS(NS.pkg, 'part')).find(
    (p) => p.getAttributeNS(NS.pkg, 'name') === partName
  );
  if (!part) throw new Error(`part ${partName} not found`);
  const xmlData = part.getElementsByTagNameNS(NS.pkg, 'xmlData')[0];
  if (!xmlData) throw new Error(`part ${partName} has no xmlData element`);
  return firstElement(xmlData);
}

describe('buildListOoxml', () => {
  it('produces a well-formed flat-OPC XML package', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'one' }] }];
    const doc = parseXml(buildListOoxml(items));
    expect(doc.documentElement.localName).toBe('package');
    expect(doc.documentElement.namespaceURI).toBe(NS.pkg);
  });

  it('includes all five required parts', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'one' }] }];
    const doc = parseXml(buildListOoxml(items));
    const names = Array.from(doc.getElementsByTagNameNS(NS.pkg, 'part')).map((p) => p.getAttributeNS(NS.pkg, 'name'));
    expect(names).toEqual(
      expect.arrayContaining(['/_rels/.rels', '/word/document.xml', '/word/_rels/document.xml.rels', '/word/numbering.xml', '/word/styles.xml'])
    );
  });

  it('one <w:p> per item, each with pStyle ListParagraph and a numPr pointing at numId 1', () => {
    const items: OoxmlListItem[] = [
      { level: 0, ordered: false, inlines: [{ text: 'a' }] },
      { level: 1, ordered: false, inlines: [{ text: 'b' }] }
    ];
    const doc = parseXml(buildListOoxml(items));
    const body = partXmlData(doc, '/word/document.xml');
    const paragraphs = body.getElementsByTagNameNS(NS.w, 'p');
    expect(paragraphs.length).toBe(2);

    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const pStyle = p.getElementsByTagNameNS(NS.w, 'pStyle')[0];
      expect(pStyle.getAttributeNS(NS.w, 'val')).toBe('ListParagraph');
      const numId = p.getElementsByTagNameNS(NS.w, 'numId')[0];
      expect(numId.getAttributeNS(NS.w, 'val')).toBe('1');
      const ilvl = p.getElementsByTagNameNS(NS.w, 'ilvl')[0];
      expect(ilvl.getAttributeNS(NS.w, 'val')).toBe(String(items[i].level));
    }
  });

  it('XML-escapes special characters in run text', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'A & B <tag> "quoted" \'x\'' }] }];
    const xml = buildListOoxml(items);
    // The raw output must not contain a literal unescaped ampersand/angle bracket outside of markup.
    expect(xml).toContain('A &amp; B &lt;tag&gt; &quot;quoted&quot; &apos;x&apos;');
    const doc = parseXml(xml);
    const body = partXmlData(doc, '/word/document.xml');
    const t = body.getElementsByTagNameNS(NS.w, 't')[0];
    expect(t.textContent).toBe('A & B <tag> "quoted" \'x\'');
  });

  it('escapeXml handles all five predefined entities', () => {
    expect(escapeXml(`&<>"'`)).toBe('&amp;&lt;&gt;&quot;&apos;');
  });

  it('emits <w:b/> and <w:i/> only for bold/italic runs, and plain runs for neither', () => {
    const items: OoxmlListItem[] = [
      {
        level: 0,
        ordered: false,
        inlines: [
          { text: 'bold', bold: true },
          { text: 'italic', italic: true },
          { text: 'both', bold: true, italic: true },
          { text: 'plain' }
        ]
      }
    ];
    const doc = parseXml(buildListOoxml(items));
    const body = partXmlData(doc, '/word/document.xml');
    const runs = Array.from(body.getElementsByTagNameNS(NS.w, 'r'));
    expect(runs).toHaveLength(4);

    const hasTag = (r: Element, tag: string) => r.getElementsByTagNameNS(NS.w, tag).length > 0;
    expect(hasTag(runs[0], 'b')).toBe(true);
    expect(hasTag(runs[0], 'i')).toBe(false);
    expect(hasTag(runs[1], 'i')).toBe(true);
    expect(hasTag(runs[1], 'b')).toBe(false);
    expect(hasTag(runs[2], 'b')).toBe(true);
    expect(hasTag(runs[2], 'i')).toBe(true);
    expect(hasTag(runs[3], 'b')).toBe(false);
    expect(hasTag(runs[3], 'i')).toBe(false);
  });

  it('splits a "\\n" inline into <w:br/>-separated text within one run', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'line one\nline two' }] }];
    const doc = parseXml(buildListOoxml(items));
    const body = partXmlData(doc, '/word/document.xml');
    const run = body.getElementsByTagNameNS(NS.w, 'r')[0];
    expect(run.getElementsByTagNameNS(NS.w, 'br').length).toBe(1);
    const texts = Array.from(run.getElementsByTagNameNS(NS.w, 't')).map((t) => t.textContent);
    expect(texts).toEqual(['line one', 'line two']);
  });

  it('numbering.xml uses decimal "%N." format for an ordered list level', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: true, inlines: [{ text: 'one' }] }];
    const doc = parseXml(buildListOoxml(items));
    const numbering = partXmlData(doc, '/word/numbering.xml');
    const lvl0 = Array.from(numbering.getElementsByTagNameNS(NS.w, 'lvl')).find((l) => l.getAttributeNS(NS.w, 'ilvl') === '0')!;
    expect(lvl0.getElementsByTagNameNS(NS.w, 'numFmt')[0].getAttributeNS(NS.w, 'val')).toBe('decimal');
    expect(lvl0.getElementsByTagNameNS(NS.w, 'lvlText')[0].getAttributeNS(NS.w, 'val')).toBe('%1.');
  });

  it('numbering.xml uses a bullet format for an unordered list level, with Word\'s default glyph/font', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'one' }] }];
    const doc = parseXml(buildListOoxml(items));
    const numbering = partXmlData(doc, '/word/numbering.xml');
    const lvl0 = Array.from(numbering.getElementsByTagNameNS(NS.w, 'lvl')).find((l) => l.getAttributeNS(NS.w, 'ilvl') === '0')!;
    expect(lvl0.getElementsByTagNameNS(NS.w, 'numFmt')[0].getAttributeNS(NS.w, 'val')).toBe('bullet');
    expect(lvl0.getElementsByTagNameNS(NS.w, 'lvlText')[0].getAttributeNS(NS.w, 'val')).toBe('');
    const rFonts = lvl0.getElementsByTagNameNS(NS.w, 'rFonts')[0];
    expect(rFonts.getAttributeNS(NS.w, 'ascii')).toBe('Symbol');
  });

  it('defines all 9 numbering levels even when only level 0 is used', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'one' }] }];
    const doc = parseXml(buildListOoxml(items));
    const numbering = partXmlData(doc, '/word/numbering.xml');
    expect(numbering.getElementsByTagNameNS(NS.w, 'lvl').length).toBe(9);
  });

  it('increases left indent per level', () => {
    const items: OoxmlListItem[] = [{ level: 2, ordered: false, inlines: [{ text: 'deep' }] }];
    const doc = parseXml(buildListOoxml(items));
    const numbering = partXmlData(doc, '/word/numbering.xml');
    const lvl2 = Array.from(numbering.getElementsByTagNameNS(NS.w, 'lvl')).find((l) => l.getAttributeNS(NS.w, 'ilvl') === '2')!;
    const ind = lvl2.getElementsByTagNameNS(NS.w, 'ind')[0];
    expect(Number(ind.getAttributeNS(NS.w, 'left'))).toBe(720 + 720 * 2);
  });

  it('defines a ListParagraph style based on Normal in styles.xml', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [{ text: 'one' }] }];
    const doc = parseXml(buildListOoxml(items));
    const styles = partXmlData(doc, '/word/styles.xml');
    const style = Array.from(styles.getElementsByTagNameNS(NS.w, 'style')).find(
      (s) => s.getAttributeNS(NS.w, 'styleId') === 'ListParagraph'
    )!;
    expect(style).toBeTruthy();
    expect(style.getElementsByTagNameNS(NS.w, 'basedOn')[0].getAttributeNS(NS.w, 'val')).toBe('Normal');
  });

  it('handles an item with no inlines (blank list item) without producing invalid XML', () => {
    const items: OoxmlListItem[] = [{ level: 0, ordered: false, inlines: [] }];
    expect(() => parseXml(buildListOoxml(items))).not.toThrow();
  });
});
