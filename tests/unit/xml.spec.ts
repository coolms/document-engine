import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { XmlDocument, XmlElement, escapeAttribute, escapeText, unescapeXml } from '../../src/ooxml/xml.js';

const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), '../fixtures/ooxml');

function fixture(file: string): string {
    return readFileSync(join(FIXTURES, file), 'utf8');
}

function parse(source: string): XmlDocument {
    return XmlDocument.parse(source);
}

describe('XmlDocument', () => {
    describe('byte-exact round trip', () => {
        // The fixtures are REAL parts: one document authored in Word (23
        // namespace declarations, mc:Ignorable, revision ids) and one produced
        // by PHPWord (Cyrillic, xml:space="preserve"). Hand-written XML would
        // only prove the parser handles the XML I thought to write.
        const files = readdirSync(FIXTURES).filter((name) => name.endsWith('.xml'));

        it('found the fixtures', () => {
            // Otherwise an empty directory would make the loop below vacuous
            // and the whole suite would pass having tested nothing.
            expect(files.length).toBeGreaterThanOrEqual(7);
        });

        for (const file of files) {
            it(`reproduces ${file} byte for byte`, () => {
                const source = fixture(file);

                expect(parse(source).toString()).toBe(source);
            });
        }
    });

    describe('surgery', () => {
        it('changes ONLY the bytes under the node that was edited', () => {
            // This is the whole thesis of not using an XML library. If editing
            // one paragraph rewrites the file, then charts, VML, revision marks
            // and everything else we do not model are at risk on every save.
            const source = fixture('word-authored.document.xml');
            const document = parse(source);

            const target = document.root.descendants('w:t')[3]!;
            const before = target.toString();
            target.setText('REPLACED');

            const output = document.toString();
            const originalAt = source.indexOf(before);
            expect(originalAt).toBeGreaterThan(0);

            // Everything before the edited element, and everything after it,
            // must be identical — compared against the ORIGINAL source, not
            // against another serialisation of the same tree.
            expect(output.slice(0, originalAt)).toBe(source.slice(0, originalAt));
            expect(output.slice(originalAt + '<w:t xml:space="preserve">REPLACED</w:t>'.length))
                .toBe(source.slice(originalAt + before.length));
        });

        it('leaves untouched SIBLINGS on their original bytes', () => {
            const source = fixture('phpword.document.xml');
            const document = parse(source);
            const runs = document.root.descendants('w:t');

            const untouched = runs[1]!.toString();
            runs[0]!.setText('changed');

            expect(runs[1]!.toString()).toBe(untouched);
        });

        it('rewrites an ancestor once, however deep the change', () => {
            const source = '<a><b><c><d>x</d></c></b><e>keep</e></a>';
            const document = parse(source);

            document.root.descendants('d')[0]!.setText('y');

            expect(document.toString()).toBe('<a><b><c><d>y</d></c></b><e>keep</e></a>');
        });

        it('keeps an attribute in its original position when changed', () => {
            // Appending it instead would turn a one-value edit into a rewritten
            // element in every diff of the resulting file.
            const document = parse('<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>');

            document.root.setAttribute('w:h', '20000');

            expect(document.toString()).toBe('<w:pgSz w:w="12240" w:h="20000" w:orient="portrait"/>');
        });

        it('does not touch the file when a value is set to what it already was', () => {
            // The odd spacing and single quotes are the point: they are legal,
            // they are not what the serialiser would produce, and so they only
            // survive if the element was never invalidated. With a value that
            // did not change, it must not be.
            const source = "<a  x='1' ><b/></a>";
            const document = parse(source);

            document.root.setAttribute('x', '1');

            expect(document.toString()).toBe(source);
        });

        it('grows a closing tag when a self-closed element is given text', () => {
            const document = parse('<w:p><w:t/></w:p>');

            document.root.element('w:t')!.setText('hello');

            expect(document.toString()).toBe('<w:p><w:t>hello</w:t></w:p>');
        });

        it('switches a single-quoted attribute to double quotes when written', () => {
            // The escaping protects `"` and not `'`, so a value holding an
            // apostrophe would end the attribute early and produce a file that
            // no longer parses.
            const document = parse("<a b='x'/>");

            document.root.setAttribute('b', "it's");

            expect(document.toString()).toBe('<a b="it\'s"/>');
            expect(parse(document.toString()).root.attribute('b')).toBe("it's");
        });

        it('escapes what it writes, and reads it back unchanged', () => {
            const document = parse('<w:t>x</w:t>');
            const awkward = 'a < b & c > d "quoted"   конец';

            document.root.setText(awkward);

            expect(document.toString()).not.toContain('a < b');
            expect(parse(document.toString()).root.text).toBe(awkward);
        });
    });

    describe('reading', () => {
        it('returns paragraphs in reading order', () => {
            // Layout consumes them in this order; a set or a map would not do.
            const document = parse(fixture('phpword.document.xml'));
            const paragraphs = document.root.descendants('w:p');

            expect(paragraphs.length).toBeGreaterThan(1);
            expect(paragraphs.map((p) => p.text)).toEqual([
                'Договор аренды',
                'Page one of the lease agreement.',
                // A paragraph holding nothing but <w:br w:type="page"/>. It has
                // no text and is not noise: it is where the document says a
                // page ends, and dropping it would merge three pages into one.
                '',
                'Условия',
                'Page two: terms and conditions.',
                '',
                'Page three: signatures.',
            ]);
        });

        it('sees structure that reading the text alone would miss', () => {
            // Two of the seven paragraphs above are empty because their whole
            // content is a page break, and one carries bookmarks around its
            // run. A text-only view of this document reports five paragraphs
            // and no page breaks — which is how "3 pages on 1 page" happens.
            const document = parse(fixture('phpword.document.xml'));

            const breaks = document.root.descendants('w:br')
                .filter((element) => 'page' === element.attribute('w:type'));
            expect(breaks.length).toBe(2);
            expect(document.root.descendants('w:bookmarkStart').length).toBe(2);
        });

        it('reads an element\'s text from ALL its descendants', () => {
            // A paragraph's text lives two levels down, inside its runs. Direct
            // children only would report every paragraph as empty.
            const document = parse('<w:p><w:r><w:t>one </w:t></w:r><w:r><w:t>two</w:t></w:r></w:p>');

            expect(document.root.text).toBe('one two');
        });

        it('separates direct children from descendants', () => {
            const document = parse('<w:body><w:p><w:p/></w:p><w:p/></w:body>');

            expect(document.root.elements('w:p').length).toBe(2);
            expect(document.root.descendants('w:p').length).toBe(3);
        });

        it('reads the page geometry a section declares', () => {
            const document = parse(fixture('word-authored.document.xml'));
            const size = document.root.descendants('w:pgSz')[0]!;

            // US Letter in twips: 8.5in x 11in at 1440 per inch.
            expect(size.attribute('w:w')).toBe('12240');
            expect(size.attribute('w:h')).toBe('15840');
            expect(size.attribute('w:nonesuch')).toBeNull();
        });

        it('ignores comments and processing instructions in text', () => {
            const document = parse('<a>x<!-- note -->y<?pi go?>z</a>');

            expect(document.root.text).toBe('xyz');
            expect(document.toString()).toBe('<a>x<!-- note -->y<?pi go?>z</a>');
        });

        it('treats CDATA as text', () => {
            const document = parse('<a><![CDATA[<not markup> & raw]]></a>');

            expect(document.root.text).toBe('<not markup> & raw');
            expect(document.toString()).toBe('<a><![CDATA[<not markup> & raw]]></a>');
        });
    });

    describe('entities', () => {
        it('resolves the five XML predefines and numeric references', () => {
            expect(unescapeXml('&amp;&lt;&gt;&quot;&apos;')).toBe('&<>"\'');
            expect(unescapeXml('&#1057;&#x422;')).toBe('СТ');
        });

        it('refuses an entity XML does not define', () => {
            // &nbsp; is HTML, not XML. Passing it through would put literal
            // "&nbsp;" in a document; resolving it would invent a character the
            // file never declared.
            expect(() => unescapeXml('a&nbsp;b')).toThrow(/undefined entity/i);
        });

        it('round-trips text through escape and back', () => {
            const awkward = '< & > " \' \t\n';

            expect(unescapeXml(escapeText(awkward))).toBe('< & > " \' \t\n');
            expect(unescapeXml(escapeAttribute(awkward))).toBe(awkward);
        });

        it('escapes tab and newline in ATTRIBUTES, where XML would normalise them', () => {
            // An attribute value's literal whitespace is normalised to spaces on
            // the next read, so a value would silently change.
            expect(escapeAttribute('a\tb\nc')).toBe('a&#9;b&#10;c');
            expect(escapeText('a\tb\nc')).toBe('a\tb\nc');
        });
    });

    describe('malformed input', () => {
        it.each([
            ['<a>', /unclosed <a>/],
            ['<a></b>', /<\/b> closes <a>/],
            ['<a', /unterminated start tag/],
            ['<a b>', /attribute b has no value/],
            ['<a b=c/>', /attribute b is not quoted/],
            ['<a b="x/>', /unterminated value/],
            ['not xml at all', /expected a root element/],
            ['<!-- unterminated', /unterminated comment/],
        ])('rejects %j rather than guessing', (source, message) => {
            expect(() => parse(source)).toThrow(message);
        });

        it('reports where the problem is', () => {
            // "Invalid XML" about a 48000-character single-line part is not a
            // diagnosis.
            expect(() => parse('<a>\n  <b></c>')).toThrow(/line 2, column 9/);
        });
    });

    describe('structure', () => {
        it('preserves the XML declaration and whatever surrounds the root', () => {
            const source = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a/>\n';

            expect(parse(source).toString()).toBe(source);
        });

        it('keeps a self-closing element self-closing', () => {
            // Word writes <w:b/> and rewriting it as <w:b></w:b> is valid XML
            // that produces a needlessly different file.
            expect(parse('<w:rPr><w:b/><w:i></w:i></w:rPr>').toString()).toBe('<w:rPr><w:b/><w:i></w:i></w:rPr>');
        });

        it('exposes the root element by name', () => {
            const root: XmlElement = parse(fixture('word-authored.document.xml')).root;

            expect(root.name).toBe('w:document');
        });

        it('removes an attribute without disturbing the others', () => {
            const document = parse('<a x="1" y="2" z="3"/>');

            document.root.removeAttribute('y');

            expect(document.toString()).toBe('<a x="1" z="3"/>');
        });
    });
});
