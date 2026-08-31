import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { FontCatalogue, type FontManifest } from '../../src/word/font-catalogue.js';
import {
    readWordDocument,
    type FurnitureVariant,
    type WordDocument,
} from '../../src/word/document-reader.js';
import {
    DEFAULT_TAB_PX,
    isFloatingBox,
    isTable,
    layoutPages,
    type Block,
    type Paragraph,
} from '../../src/layout/page-layout.js';
import { eighthPointsToPx, emuToPx, halfPointsToPx, twipsToPx } from '../../src/ooxml/units.js';
import { OBJECT_REPLACEMENT } from '../../src/layout/image.js';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');
const FIXTURES = join(ROOT, 'tests/fixtures/ooxml');
const FONT_DIR = join(ROOT, 'assets/fonts');

const MANIFEST = JSON.parse(readFileSync(join(FONT_DIR, 'fonts.manifest.json'), 'utf8')) as FontManifest;
const FONTS = FontCatalogue.load(MANIFEST, (file) => new Uint8Array(readFileSync(join(FONT_DIR, file))));

/**
 * The block at `index`, as a paragraph.
 *
 * Throws rather than casts: a document holds paragraphs AND tables now, and a
 * test that reaches for `.runs` on a table should fail loudly rather than read
 * undefined and pass.
 */
function para(document: { paragraphs: readonly Block[] }, index: number): Paragraph {
    const block = document.paragraphs[index];
    if (undefined === block || isTable(block)) {
        throw new Error(`block ${index} is not a paragraph`);
    }

    return block;
}

function fixture(file: string): string {
    return readFileSync(join(FIXTURES, file), 'utf8');
}

function read(documentXml: string, stylesXml?: string) {
    return readWordDocument(
        undefined === stylesXml
            ? { documentXml, fonts: FONTS }
            : { documentXml, stylesXml, fonts: FONTS },
    );
}

function readNumbered(documentXml: string, numberingXml: string) {
    return readWordDocument({ documentXml, numberingXml, fonts: FONTS });
}

function readWithFurniture(
    documentXml: string,
    headers?: Partial<Record<'default' | 'first' | 'even', string>>,
    footers?: Partial<Record<'default' | 'first' | 'even', string>>,
) {
    return readWordDocument({
        documentXml,
        fonts: FONTS,
        ...(undefined === headers ? {} : { headers }),
        ...(undefined === footers ? {} : { footers }),
    });
}

/** A minimal document body, so a test can state one construct and nothing else. */
function doc(body: string): string {
    return `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
        + `<w:body>${body}</w:body></w:document>`;
}

function paragraph(text: string, pPr = ''): string {
    return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

describe('readWordDocument', () => {
    describe('a real document, end to end', () => {
        it('paginates the three-page lease agreement as THREE pages', () => {
            // The document says so itself: its paragraphs read "Page one",
            // "Page two", "Page three", separated by explicit page breaks. This
            // is the original complaint — "3 pages on 1 page" — asserted
            // against a real file rather than a fixture written to pass.
            const document = read(fixture('phpword.document.xml'), fixture('phpword.styles.xml'));
            const pages = layoutPages(document.paragraphs, document.geometry);

            expect(pages.length).toBe(3);

            const textOf = (index: number): string =>
                pages[index]!.lines.flatMap((line) => line.line.pieces.map((piece) => piece.text)).join('');

            expect(textOf(0)).toContain('Page one of the lease agreement.');
            expect(textOf(1)).toContain('Page two: terms and conditions.');
            expect(textOf(2)).toContain('Page three: signatures.');

            // And the pages must not bleed into one another.
            expect(textOf(0)).not.toContain('Page two');
            expect(textOf(1)).not.toContain('Page three');
        });

        it('reads that document\'s geometry from its own w:sectPr, LANDSCAPE and all', () => {
            const document = read(fixture('phpword.document.xml'), fixture('phpword.styles.xml'));

            // This file declares w:orient="landscape" with w:w="16838". OOXML
            // stores the dimensions ALREADY swapped, so honouring the
            // orientation flag on top of that would rotate the page back to
            // portrait and repaginate the whole document.
            expect(document.geometry.widthPx).toBe(twipsToPx(16838));
            expect(document.geometry.heightPx).toBe(twipsToPx(11906));
            expect(document.geometry.widthPx).toBeGreaterThan(document.geometry.heightPx);
        });

        it('substitutes the Arial its docDefaults ask for, and says so', () => {
            // The whole document is Arial, declared once in docDefaults. If the
            // substitution did not happen the runs would silently fall back to
            // the default family, which is a different width.
            const document = read(fixture('phpword.document.xml'), fixture('phpword.styles.xml'));

            expect(document.diagnostics).toContainEqual({
                kind: 'font-substituted',
                detail: '"Arial" is not available; using the metric-compatible Liberation Sans',
            });
        });

        it('reads a Word-authored document, tables included', () => {
            const document = read(
                fixture('word-authored.document.xml'),
                fixture('word-authored.styles.xml'),
            );

            // 160 paragraphs in the file: 100 at body level and 60 inside its
            // four tables. The blocks are those 100 plus the 4 tables, and the
            // tables are now laid out rather than reported as skipped.
            const tables = document.paragraphs.filter(isTable);
            expect(document.paragraphs.length).toBe(104);
            expect(tables.length).toBe(4);
            expect(tables.reduce((sum, table) =>
                sum + table.rows.reduce((cells, row) =>
                    cells + row.cells.reduce((paragraphs, cell) => paragraphs + cell.paragraphs.length, 0), 0), 0))
                .toBe(60);
            expect(document.diagnostics.map((d) => d.kind)).not.toContain('unsupported-block');

            expect(document.geometry.widthPx).toBe(twipsToPx(12240)); // US Letter
            expect(document.geometry.marginLeftPx).toBe(96); // one inch

            // Its styles are all basedOn="Normal" and it HAS no Normal style.
            // The dangling reference must not throw and must not be reported as
            // an unknown style either — the paragraphs name styles that exist.
            const unknown = document.diagnostics.filter((d) => 'unknown-paragraph-style' === d.kind);
            expect(unknown).toEqual([]);
        });

        it('lays that document out onto many pages, all within the page box', () => {
            const document = read(
                fixture('word-authored.document.xml'),
                fixture('word-authored.styles.xml'),
            );
            const pages = layoutPages(document.paragraphs, document.geometry);

            expect(pages.length).toBeGreaterThan(1);
            for (const page of pages) {
                for (const line of page.lines) {
                    expect(line.yPx).toBeGreaterThanOrEqual(document.geometry.marginTopPx);
                    expect(line.yPx + line.heightPx)
                        .toBeLessThanOrEqual(document.geometry.heightPx - document.geometry.marginBottomPx);
                }
            }
        });
    });

    describe('page breaks inside a paragraph', () => {
        it('splits the paragraph and starts the remainder on a new page', () => {
            const document = read(doc(
                '<w:p><w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>',
            ));

            expect(document.paragraphs.length).toBe(2);
            expect(para(document, 0).style.pageBreakBefore).toBe(false);
            expect(para(document, 1).style.pageBreakBefore).toBe(true);
            expect(para(document, 1).runs[0]!.text).toBe('after');
        });

        it('does not repeat the paragraph\'s spacing at the seam', () => {
            // Space before belongs to the start of the paragraph and space
            // after to its end. Repeating either at a break would push the
            // second half down a gap the document never asked for.
            const document = read(doc(
                '<w:p><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr>'
                + '<w:r><w:t>a</w:t><w:br w:type="page"/><w:t>b</w:t></w:r></w:p>',
            ));

            const first = para(document, 0);
            const second = para(document, 1);
            expect(first.style.spaceBeforePx).toBe(twipsToPx(240));
            expect(first.style.spaceAfterPx).toBe(0);
            expect(second.style.spaceBeforePx).toBe(0);
            expect(second.style.spaceAfterPx).toBe(twipsToPx(240));
        });

        it('treats a plain w:br as a line break, not a page break', () => {
            const document = read(doc('<w:p><w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r></w:p>'));

            expect(document.paragraphs.length).toBe(1);
            expect(para(document, 0).runs.map((run) => run.text).join('')).toBe('a\nb');
        });
    });

    describe('style resolution', () => {
        const styles = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>
                <w:pPrDefault><w:pPr><w:spacing w:after="100"/></w:pPr></w:pPrDefault></w:docDefaults>
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>
                <w:pPr><w:ind w:left="60"/></w:pPr></w:style>
            <w:style w:type="paragraph" w:styleId="Base"><w:basedOn w:val="Normal"/>
                <w:pPr><w:spacing w:before="200" w:after="200"/><w:keepLines/></w:pPr>
                <w:rPr><w:sz w:val="40"/><w:b/></w:rPr></w:style>
            <w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/>
                <w:pPr><w:spacing w:after="300"/></w:pPr></w:style>
        </w:styles>`;

        it('inherits through the basedOn chain', () => {
            const document = read(
                doc(paragraph('x', '<w:pPr><w:pStyle w:val="Derived"/></w:pPr>')),
                styles,
            );
            const style = para(document, 0).style;

            // spaceBefore comes from Base, spaceAfter is overridden by Derived,
            // keepLines is inherited from Base through Derived.
            expect(style.spaceBeforePx).toBe(twipsToPx(200));
            expect(style.spaceAfterPx).toBe(twipsToPx(300));
            expect(style.keepLinesTogether).toBe(true);
        });

        it('lets direct formatting beat the style', () => {
            const document = read(
                doc(paragraph('x', '<w:pPr><w:pStyle w:val="Derived"/><w:spacing w:after="500"/></w:pPr>')),
                styles,
            );

            expect(para(document, 0).style.spaceAfterPx).toBe(twipsToPx(500));
        });

        it('falls back to docDefaults when no style applies', () => {
            const document = read(doc(paragraph('x')), styles);

            expect(para(document, 0).style.spaceAfterPx).toBe(twipsToPx(100));
        });

        it('ignores a conditional format that says no w:type', () => {
            // `w:type` is required on `w:tblStylePr`, so a file without it is
            // malformed and there is nothing measured to follow. Reading it as
            // `firstRow` — the tempting guess, since that is the common one —
            // would shade a header row the document never asked to shade.
            const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
            const document = read(
                doc('<w:tbl><w:tblPr><w:tblStyle w:val="T"/></w:tblPr>'
                    + '<w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>'
                    + '<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
                `<w:styles ${NS}><w:docDefaults/>`
                + '<w:style w:type="table" w:styleId="T"><w:name w:val="T"/>'
                + '<w:tblStylePr><w:tcPr><w:shd w:val="clear" w:fill="C0C0C0"/></w:tcPr>'
                + '</w:tblStylePr></w:style></w:styles>',
            );
            const table = document.paragraphs[0]!;

            expect(isTable(table)).toBe(true);
            expect(isTable(table) ? table.rows[0]!.cells[0]!.shadingFill : 'x').toBeUndefined();
        });

        it('spaces a style named "heading 2" and not one merely CONTAINING it', () => {
            // The built-in heading spacing keys on the whole name,
            // `heading 1` through `heading 9`. Word ships a built-in called
            // `TOC Heading`, and a contents heading pushed 12pt down the page
            // is a real document breaking on a real file — so the negative case
            // here is that name rather than an invented one.
            const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
            const named = (id: string, name: string): string =>
                `<w:style w:type="paragraph" w:styleId="${id}">`
                + `<w:name w:val="${name}"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>`;
            const document = read(
                doc(paragraph('a', '<w:pPr><w:pStyle w:val="H2"/></w:pPr>')
                    + paragraph('b', '<w:pPr><w:pStyle w:val="TocH"/></w:pPr>')
                    // The format defines nine heading levels; a tenth is a
                    // style of the author's own that happens to read like one.
                    + paragraph('c', '<w:pPr><w:pStyle w:val="H10"/></w:pPr>')),
                `<w:styles ${NS}><w:docDefaults/>`
                + named('H2', 'heading 2') + named('TocH', 'TOC Heading')
                + named('H10', 'heading 10') + '</w:styles>',
            );

            expect(para(document, 0).style.spaceBeforePx).toBe(twipsToPx(240));
            expect(para(document, 0).style.spaceAfterPx).toBe(twipsToPx(120));
            expect(para(document, 1).style.spaceBeforePx).toBe(0);
            expect(para(document, 1).style.spaceAfterPx).toBe(0);
            expect(para(document, 2).style.spaceBeforePx).toBe(0);
        });

        it('takes the SIZE from docDefaults, not the fallback for a silent one', () => {
            // What every real Word document relies on: a size stated once in
            // `rPrDefault` and never again. The fallback a styles part gets
            // when it states nothing has to stand aside for it.
            //
            // THIRTEEN points, deliberately — the shared fixture above says ten,
            // which is the fallback's own number, so a fallback written to
            // overwrite rather than to fill a gap would pass against it and
            // change every real document's body size on the way past.
            const thirteen = `<w:styles xmlns:w="http://schemas.openxmlformats.org/`
                + `wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>`
                + `<w:sz w:val="26"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;
            const document = read(doc(paragraph('x')), thirteen);

            expect(para(document, 0).runs[0]!.sizePx).toBe(halfPointsToPx(26));
        });

        it('applies the w:default="1" style to a paragraph that names none', () => {
            // Almost every paragraph in a real document has no w:pStyle and is
            // formatted entirely by Normal. Ignoring the default style leaves
            // all of them with document defaults only.
            const document = read(doc(paragraph('x')), styles);

            expect(para(document, 0).style.indentLeftPx).toBe(twipsToPx(60));
        });

        it('takes a run\'s size from the PARAGRAPH style, where headings keep it', () => {
            // Heading1 in the Word fixture carries nothing but a w:rPr. A
            // reader that only consulted character styles would render every
            // heading at body size.
            const document = read(
                doc(paragraph('x', '<w:pPr><w:pStyle w:val="Base"/></w:pPr>')),
                styles,
            );

            expect(para(document, 0).runs[0]!.sizePx).toBe(twipsToPx(40 * 10)); // 20pt
        });

        it('survives a basedOn that names a style the file does not contain', () => {
            const orphan = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/>
                    <w:rPr><w:sz w:val="32"/></w:rPr></w:style></w:styles>`;

            const document = read(
                doc(paragraph('x', '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>')),
                orphan,
            );

            expect(para(document, 0).runs[0]!.sizePx).toBe(twipsToPx(16 * 20)); // 16pt
        });

        it('does not hang on a basedOn cycle', () => {
            const cyclic = `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="B"/></w:style>
                <w:style w:type="paragraph" w:styleId="B"><w:basedOn w:val="A"/>
                    <w:rPr><w:sz w:val="30"/></w:rPr></w:style></w:styles>`;

            const document = read(doc(paragraph('x', '<w:pPr><w:pStyle w:val="A"/></w:pPr>')), cyclic);

            expect(para(document, 0).runs[0]!.sizePx).toBe(twipsToPx(15 * 20)); // 15pt
        });

        it('reports a w:pStyle that names nothing', () => {
            const document = read(doc(paragraph('x', '<w:pPr><w:pStyle w:val="Nope"/></w:pPr>')), styles);

            expect(document.diagnostics).toContainEqual({
                kind: 'unknown-paragraph-style',
                detail: 'w:pStyle "Nope" is not defined in styles.xml',
            });
        });
    });

    describe('paragraph properties', () => {
        it('reads a hanging indent as a NEGATIVE first-line indent', () => {
            // w:hanging is what every bulleted list uses. Reading only
            // w:firstLine turns it into no indent at all.
            const document = read(doc(paragraph('x', '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>')));
            const style = para(document, 0).style;

            expect(style.indentLeftPx).toBe(twipsToPx(720));
            expect(style.indentFirstLinePx).toBe(twipsToPx(-360));
        });

        it('accepts w:start and w:end as well as w:left and w:right', () => {
            const document = read(doc(paragraph('x', '<w:pPr><w:ind w:start="720" w:end="360"/></w:pPr>')));
            const style = para(document, 0).style;

            expect(style.indentLeftPx).toBe(twipsToPx(720));
            expect(style.indentRightPx).toBe(twipsToPx(360));
        });

        it('reads an explicit page break before a paragraph', () => {
            const document = read(doc(paragraph('x', '<w:pPr><w:pageBreakBefore/></w:pPr>')));

            expect(para(document, 0).style.pageBreakBefore).toBe(true);
        });

        it('honours a toggle turned explicitly OFF', () => {
            // <w:widowControl w:val="0"/> is how a paragraph escapes a setting
            // that is on by default. Reading presence alone makes it
            // impossible to turn anything off.
            const on = read(doc(paragraph('x', '<w:pPr><w:widowControl/></w:pPr>')));
            const off = read(doc(paragraph('x', '<w:pPr><w:widowControl w:val="0"/></w:pPr>')));

            expect(para(on, 0).style.widowControl).toBe(true);
            expect(para(off, 0).style.widowControl).toBe(false);
        });
    });

    describe('line spacing', () => {
        const withSpacing = (spacing: string): number | undefined =>
            para(read(doc(paragraph('x', `<w:pPr>${spacing}</w:pPr>`))), 0).style.lineHeightPx;

        it('reads w:line as 240ths of a line when the rule is auto', () => {
            // 360 is one-and-a-half spacing, not 360 twips.
            const single = para(read(doc(paragraph('x'))), 0);
            const natural = single.runs[0]!.font.naturalLineHeight(single.runs[0]!.sizePx);

            expect(withSpacing('<w:spacing w:line="360" w:lineRule="auto"/>')).toBeCloseTo(natural * 1.5, 9);
            expect(withSpacing('<w:spacing w:line="480" w:lineRule="auto"/>')).toBeCloseTo(natural * 2, 9);
        });

        it('treats a missing lineRule as auto, as Word does', () => {
            expect(withSpacing('<w:spacing w:line="360"/>'))
                .toBe(withSpacing('<w:spacing w:line="360" w:lineRule="auto"/>'));
        });

        it('reads w:line as TWIPS when the rule is exact', () => {
            expect(withSpacing('<w:spacing w:line="360" w:lineRule="exact"/>')).toBe(twipsToPx(360));
        });

        it('treats atLeast as a floor, not a fixed height', () => {
            const single = para(read(doc(paragraph('x'))), 0);
            const natural = single.runs[0]!.font.naturalLineHeight(single.runs[0]!.sizePx);

            // Below the natural height it must have no effect at all.
            expect(withSpacing('<w:spacing w:line="20" w:lineRule="atLeast"/>')).toBeCloseTo(natural, 9);
            expect(withSpacing('<w:spacing w:line="1000" w:lineRule="atLeast"/>')).toBe(twipsToPx(1000));
        });

        it('leaves line height to the font when the document says nothing', () => {
            expect(withSpacing('')).toBeUndefined();
        });
    });

    describe('runs and content', () => {
        it('keeps the text of a hyperlink, which wraps its runs', () => {
            const document = read(doc(
                '<w:p><w:hyperlink r:id="rId1"><w:r><w:t>linked text</w:t></w:r></w:hyperlink></w:p>',
            ));

            expect(para(document, 0).runs.map((run) => run.text).join('')).toBe('linked text');
        });

        it('does not draw text deleted under revision tracking', () => {
            // Word writes deleted text as w:delText, which this reader would
            // ignore anyway. The w:t case is the one that actually exercises
            // the rule — without skipping the w:del wrapper it renders, and a
            // deleted paragraph reappears in the laid-out document.
            const realistic = read(doc(
                '<w:p><w:r><w:t>kept </w:t></w:r><w:del><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>',
            ));
            const strict = read(doc(
                '<w:p><w:r><w:t>kept </w:t></w:r><w:del><w:r><w:t>gone</w:t></w:r></w:del></w:p>',
            ));

            expect(para(realistic, 0).runs.map((run) => run.text).join('')).toBe('kept ');
            expect(para(strict, 0).runs.map((run) => run.text).join('')).toBe('kept ');
        });

        it('keeps text INSERTED under revision tracking', () => {
            const document = read(doc(
                '<w:p><w:ins><w:r><w:t>added</w:t></w:r></w:ins></w:p>',
            ));

            expect(para(document, 0).runs.map((run) => run.text).join('')).toBe('added');
        });

        it('gives an empty paragraph a real height', () => {
            // Otherwise a blank line between two paragraphs takes no space and
            // every page holds more than it should.
            const document = read(doc('<w:p/>'));
            const [page] = layoutPages(document.paragraphs, document.geometry);

            expect(document.paragraphs.length).toBe(1);
            expect(page!.lines.length).toBe(1);
            expect(page!.lines[0]!.heightPx).toBeGreaterThan(0);
        });

        it('carries bold and italic into the font that gets used', () => {
            const plain = read(doc('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
            const bold = read(doc('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>'));

            const width = (d: ReturnType<typeof read>): number => {
                const run = para(d, 0).runs[0]!;

                return run.font.measureAdvance('Wagon', run.sizePx).widthPt;
            };

            // The bold face is a different file with different advances. Equal
            // widths would mean the weight never reached the font catalogue.
            expect(width(bold)).not.toBe(width(plain));
        });
    });

    describe('tables', () => {
        const grid = '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="1500"/></w:tblGrid>';

        const tableXml = (rows: string, tblPr = ''): string =>
            `<w:tbl>${tblPr}${grid}${rows}</w:tbl>`;

        const readTable = (xml: string) => {
            const block = read(doc(xml)).paragraphs[0];
            if (undefined === block || !isTable(block)) {
                throw new Error('expected a table');
            }

            return block;
        };

        it('takes its column widths from w:tblGrid when the cells declare none', () => {
            const table = readTable(tableXml('<w:tr><w:tc>' + paragraph('a') + '</w:tc></w:tr>'));

            expect(table.columnWidthsPx).toEqual([twipsToPx(3000), twipsToPx(1500)]);
        });

        it('lets w:tcW OVERRIDE a nominal grid', () => {
            // Real Word output does this: the grid says w:w="100" for every
            // column and the true width sits on each cell. Trusting the grid
            // gives columns a third of a character wide, every cell wraps to one
            // letter per line, and the document gains a page.
            const nominal = '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>';
            const cells = '<w:tr>'
                + '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4680"/></w:tcPr>' + paragraph('a') + '</w:tc>'
                + '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4680"/></w:tcPr>' + paragraph('b') + '</w:tc>'
                + '</w:tr>';

            const block = read(doc(`<w:tbl>${nominal}${cells}</w:tbl>`)).paragraphs[0];
            if (undefined === block || !isTable(block)) {
                throw new Error('expected a table');
            }

            expect(block.columnWidthsPx).toEqual([twipsToPx(4680), twipsToPx(4680)]);
        });

        it('apportions a SPANNING cell across the columns it covers', () => {
            const grid3 = '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>';
            const spanning = '<w:tr><w:tc>'
                + '<w:tcPr><w:gridSpan w:val="2"/><w:tcW w:type="dxa" w:w="4000"/></w:tcPr>'
                + paragraph('wide') + '</w:tc></w:tr>';

            const block = read(doc(`<w:tbl>${grid3}${spanning}</w:tbl>`)).paragraphs[0];
            if (undefined === block || !isTable(block)) {
                throw new Error('expected a table');
            }

            // An all-equal grid carries no ratio, so the span splits evenly.
            expect(block.columnWidthsPx).toEqual([twipsToPx(2000), twipsToPx(2000)]);
        });

        it('lets an exact single-column width beat a share of a span', () => {
            // A row of ordinary cells knows its columns better than a merged
            // cell above it does.
            const grid2 = '<w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>';
            const rows = '<w:tr><w:tc>'
                + '<w:tcPr><w:gridSpan w:val="2"/><w:tcW w:type="dxa" w:w="4000"/></w:tcPr>'
                + paragraph('title') + '</w:tc></w:tr>'
                + '<w:tr>'
                + '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="3000"/></w:tcPr>' + paragraph('a') + '</w:tc>'
                + '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="1000"/></w:tcPr>' + paragraph('b') + '</w:tc>'
                + '</w:tr>';

            const block = read(doc(`<w:tbl>${grid2}${rows}</w:tbl>`)).paragraphs[0];
            if (undefined === block || !isTable(block)) {
                throw new Error('expected a table');
            }

            expect(block.columnWidthsPx).toEqual([twipsToPx(3000), twipsToPx(1000)]);
        });

        it('reports a percentage column width instead of misreading it', () => {
            // A percentage resolves against the table's own width, which is a
            // second rule this does not implement. The grid still gives a usable
            // shape, so it falls back and says so.
            const pct = '<w:tr><w:tc><w:tcPr><w:tcW w:type="pct" w:w="2500"/></w:tcPr>'
                + paragraph('a') + '</w:tc></w:tr>';
            const document = read(doc(tableXml(pct)));

            expect(document.diagnostics.map((d) => d.detail))
                .toContain('a table column width of type "pct" fell back to the grid');
        });

        it('defaults a cell to ONE grid column', () => {
            const table = readTable(tableXml('<w:tr><w:tc>' + paragraph('a') + '</w:tc></w:tr>'));

            expect(table.rows[0]!.cells[0]!.gridSpan).toBe(1);
        });

        it('reads w:gridSpan when a cell covers several columns', () => {
            const spanning = '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'
                + paragraph('wide') + '</w:tc></w:tr>';

            expect(readTable(tableXml(spanning)).rows[0]!.cells[0]!.gridSpan).toBe(2);
        });

        it('marks a w:tblHeader row as one that repeats', () => {
            const rows = '<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc>' + paragraph('h') + '</w:tc></w:tr>'
                + '<w:tr><w:tc>' + paragraph('d') + '</w:tc></w:tr>';
            const table = readTable(tableXml(rows));

            expect(table.rows.map((row) => row.isHeader)).toEqual([true, false]);
        });

        it('reads w:cantSplit, which Word uses to keep a row whole', () => {
            const rows = '<w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc>' + paragraph('x') + '</w:tc></w:tr>';

            expect(readTable(tableXml(rows)).rows[0]!.cantSplit).toBe(true);
        });

        it('gives a table Word\'s default cell padding when it declares none', () => {
            // 115 twips either side. A table with no w:tblCellMar still has it,
            // so assuming zero makes every column a character wider than Word's.
            const table = readTable(tableXml('<w:tr><w:tc>' + paragraph('a') + '</w:tc></w:tr>'));

            expect(table.cellMarginLeftPx).toBe(twipsToPx(115));
            expect(table.cellMarginRightPx).toBe(twipsToPx(115));
            expect(table.cellMarginTopPx).toBe(0);
        });

        it('reads w:tblCellMar when the table declares it', () => {
            const tblPr = '<w:tblPr><w:tblCellMar>'
                + '<w:left w:w="200"/><w:right w:w="300"/><w:top w:w="40"/><w:bottom w:w="60"/>'
                + '</w:tblCellMar></w:tblPr>';
            const table = readTable(tableXml('<w:tr><w:tc>' + paragraph('a') + '</w:tc></w:tr>', tblPr));

            expect(table.cellMarginLeftPx).toBe(twipsToPx(200));
            expect(table.cellMarginRightPx).toBe(twipsToPx(300));
            expect(table.cellMarginTopPx).toBe(twipsToPx(40));
            expect(table.cellMarginBottomPx).toBe(twipsToPx(60));
        });

        it('reads every paragraph in every cell', () => {
            const rows = '<w:tr>'
                + '<w:tc>' + paragraph('one') + paragraph('two') + '</w:tc>'
                + '<w:tc>' + paragraph('three') + '</w:tc>'
                + '</w:tr>';
            const table = readTable(tableXml(rows));

            expect(table.rows[0]!.cells.map((cell) => cell.paragraphs.length)).toEqual([2, 1]);
            const second = table.rows[0]!.cells[0]!.paragraphs[1]!;

            expect(isTable(second) ? '' : second.runs[0]!.text).toBe('two');
        });
    });


    describe('w:cols', () => {
        const sectionWith = (cols: string): string =>
            '<w:sectPr>' + cols + '<w:pgSz w:w="12240" w:h="15840"/>'
            + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

        const columnsOf = (cols: string) => read(
            doc(paragraph('body') + sectionWith(cols)),
        ).geometry.columns;

        // 12240 - 2880 = 9360 twips of writing width.
        const px = (twips: number): number => twipsToPx(twips);

        it('divides the writing width between equal columns', () => {
            // Two columns and a 708-twip gap: (9360 - 708) / 2 each, the second
            // one gap further along than the first one ends.
            const columns = columnsOf('<w:cols w:num="2" w:space="708"/>');

            expect(columns?.length).toBe(2);
            expect(columns?.[0]?.leftPx).toBeCloseTo(px(1440), 9);
            expect(columns?.[0]?.widthPx).toBeCloseTo(px(4326), 9);
            expect(columns?.[1]?.leftPx).toBeCloseTo(px(1440 + 4326 + 708), 9);
        });

        it('takes half an inch as the gap when none is stated', () => {
            const columns = columnsOf('<w:cols w:num="2"/>');

            expect(columns?.[0]?.widthPx).toBe(px((9360 - 720) / 2));
        });

        it('divides three the same way', () => {
            const columns = columnsOf('<w:cols w:num="3" w:space="708"/>');

            expect(columns?.length).toBe(3);
            expect(columns?.[2]?.leftPx).toBeCloseTo(px(1440 + ((9360 - 1416) / 3 + 708) * 2), 9);
        });

        it('reads stated widths when the document turns equal widths OFF', () => {
            const columns = columnsOf(
                '<w:cols w:num="2" w:equalWidth="0">'
                + '<w:col w:w="7000" w:space="400"/><w:col w:w="1960"/></w:cols>',
            );

            expect(columns?.[0]?.leftPx).toBeCloseTo(px(1440), 9);
            expect(columns?.[0]?.widthPx).toBeCloseTo(px(7000), 9);
            expect(columns?.[1]?.leftPx).toBeCloseTo(px(1440 + 7400), 9);
            expect(columns?.[1]?.widthPx).toBeCloseTo(px(1960), 9);
        });

        it('ignores stated widths only where equal widths are asked for outright', () => {
            // The schema says `w:equalWidth` defaults to true, and the page
            // says otherwise: the same two columns put their second
            // at 247.10 with `"0"`, with `"false"` and with the attribute
            // ABSENT, and only `"1"` divided the width evenly, at 315.75.
            //
            // So a stated width wins unless the document asks for equal ones,
            // and the attribute is read as the toggle it is — `0`, `false` and
            // `off` all turning it off.
            const stated = '<w:col w:w="7000"/><w:col w:w="1960"/>';
            const widths = (attribute: string): (number | undefined)[] =>
                (columnsOf(`<w:cols w:num="2"${attribute}>${stated}</w:cols>`) ?? [])
                    .map((column) => column.widthPx);

            for (const attribute of ['', ' w:equalWidth="0"', ' w:equalWidth="false"',
                ' w:equalWidth="off"']) {
                expect([attribute, widths(attribute)])
                    .toEqual([attribute, [px(7000), px(1960)]]);
            }

            const even = widths(' w:equalWidth="1"');
            expect(even[0]).toBe(even[1]);
        });

        it('leaves a one-column section without any', () => {
            expect(columnsOf('<w:cols w:num="1" w:space="708"/>')).toBeUndefined();
            expect(columnsOf('')).toBeUndefined();
        });
    });


    describe('footnotes', () => {
        const NS_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
        const notesXml = (...ids: number[]): string =>
            `<w:footnotes ${NS_W}>`
            + '<w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r>'
            + '</w:p></w:footnote>'
            + ids.map((id) => `<w:footnote w:id="${id}"><w:p>`
                + '<w:r><w:footnoteRef/></w:r>'
                + `<w:r><w:t xml:space="preserve"> note ${id}</w:t></w:r>`
                + '</w:p></w:footnote>').join('')
            + '</w:footnotes>';

        const mark = (id: number): string =>
            `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

        const readNoted = (body: string, ...ids: number[]) => readWordDocument({
            documentXml: doc(body),
            fonts: FONTS,
            footnotesXml: notesXml(...ids),
        });

        it('numbers the notes by the order the BODY refers to them', () => {
            // Word writes ids in whatever order the notes were CREATED, so a
            // document edited more than once has them out of order. The mark a
            // reader sees is the position among the references.
            const document = readNoted(
                `<w:p><w:r><w:t>a</w:t></w:r>${mark(7)}</w:p>`
                + `<w:p><w:r><w:t>b</w:t></w:r>${mark(3)}</w:p>`,
                3, 7,
            );
            const marks = document.paragraphs
                .filter((block): block is Paragraph => !isTable(block))
                .flatMap((block) => block.runs)
                .filter((run) => undefined !== run.footnoteId);

            expect(marks.map((run) => [run.footnoteId, run.text])).toEqual([[7, '1'], [3, '2']]);
        });

        it('gives the note its OWN number, and no room to reserve', () => {
            // The mark that opens a note is the same number, but it is already
            // at the foot of the page: carrying an id would have the page
            // reserve room for a note it is in the middle of drawing.
            const document = readNoted(`<w:p><w:r><w:t>a</w:t></w:r>${mark(7)}</w:p>`, 7);
            const note = document.footnotes.get(7)!;
            const opening = (note[0] as Paragraph).runs[0]!;

            expect(opening.text).toBe('1');
            expect(opening.footnoteId).toBeUndefined();
        });

        it('reads no note the body never refers to', () => {
            const document = readNoted('<w:p><w:r><w:t>a</w:t></w:r></w:p>', 1, 2);

            expect(document.footnotes.size).toBe(0);
        });
    });


    describe('VML pictures', () => {
        const NS_V = 'xmlns:v="urn:schemas-microsoft-com:vml"';
        const picture = (style: string, imagedata = '<v:imagedata r:id="rIdImg"/>'): string =>
            `<w:r><w:pict><v:shape ${NS_V} style="${style}">${imagedata}`
            + '</v:shape></w:pict></w:r>';

        const readPict = (style: string, imagedata?: string) => readWordDocument({
            documentXml: doc(`<w:p>${picture(style, imagedata)}</w:p>`),
            fonts: FONTS,
            mediaById: { rIdImg: { bytes: new Uint8Array([1]), contentType: 'image/png' } },
        });

        const imageOf = (style: string) => {
            const block = readPict(style).paragraphs[0]!;

            return isTable(block) ? undefined : block.runs[0]!.image;
        };

        it('reads the picture a w:pict names', () => {
            const image = imageOf('width:144pt;height:72pt');

            expect(image?.content.contentType).toBe('image/png');
            expect(image?.widthPx).toBeCloseTo(192, 6);
            expect(image?.heightPx).toBeCloseTo(96, 6);
        });

        it('understands the units a converted document may use', () => {
            // The same picture three ways round. LibreOffice printed all three
            // identically, so all three have to mean the same thing here.
            const points = imageOf('width:144pt;height:72pt');

            expect(imageOf('width:2in;height:1in')).toEqual(points);
            expect(imageOf('width:5.08cm;height:2.54cm')).toEqual(points);
            expect(imageOf('width:50.8mm;height:25.4mm')).toEqual(points);
            // A bare number is pixels, which is what CSS says it is.
            expect(imageOf('width:192;height:96')).toEqual(points);
        });

        it('ignores a declaration that is not a length', () => {
            const image = imageOf('position:absolute;width:36pt;height:36pt;z-index:1');

            expect(image?.widthPx).toBeCloseTo(48, 6);
        });

        it('ends the text that shared the run with it', () => {
            // A run may hold text AND a picture. Without a break between them
            // the two become one piece, and the text is measured at the
            // picture's width.
            const document = readWordDocument({
                documentXml: doc('<w:p><w:r><w:t>before</w:t><w:pict>'
                    + `<v:shape ${NS_V} style="width:36pt;height:36pt">`
                    + '<v:imagedata r:id="rIdImg"/></v:shape></w:pict></w:r></w:p>'),
                fonts: FONTS,
                mediaById: { rIdImg: { bytes: new Uint8Array([1]), contentType: 'image/png' } },
            });
            const block = document.paragraphs[0]!;
            const runs = isTable(block) ? [] : block.runs;

            expect(runs.map((run) => run.text)).toEqual(['before', '\uFFFC']);
            expect(runs[0]!.image).toBeUndefined();
            expect(runs[1]!.image?.widthPx).toBe(48);
        });

        it('matches the declaration by its NAME, not by what it starts with', () => {
            // A VML shape carries `mso-width-percent` and `mso-width-relative`
            // beside its width. Taking the first declaration that merely
            // CONTAINS "width" reads one of those as the picture's size.
            const image = imageOf('mso-width-percent:0;width:36pt;height:36pt');

            expect(image?.widthPx).toBeCloseTo(48, 6);
        });

        it('reads no size at all from a unit it does not know', () => {
            // Guessing pixels would draw a picture at whatever the number
            // happened to be, which is worse than drawing it at nothing and
            // saying so through the size.
            const image = imageOf('width:36qq;height:36qq');

            expect(image?.widthPx).toBe(0);
            expect(image?.heightPx).toBe(0);
        });

        it('DRAWS a w:pict that holds no picture', () => {
            // Furniture rather than content: a rule, a border, a coloured
            // panel. Reporting it was the old behaviour and left a document
            // whose boxes are VML with none of them.
            const block = readPict('width:36pt;height:36pt', '').paragraphs[0]!;
            const runs = isTable(block) ? [] : block.runs;

            expect(runs[0]!.shape?.widthPx).toBe(48);
            expect(runs[0]!.shape?.heightPx).toBe(48);
        });

        it('fills and strokes it the way VML says to', () => {
            // On a `v:rect`, which HAS a rectangle's geometry. The colour
            // grammar is the same on any VML element; the reason it cannot be
            // shown on a bare `v:shape` is that such a shape is painted by
            // nobody — see the shapetype test below.
            const shapeOf = (attributes: string) => {
                const block = readWordDocument({
                    documentXml: doc(`<w:p><w:r><w:pict><v:rect ${NS_V} `
                        + `style="width:36pt;height:12pt" ${attributes}/></w:pict></w:r></w:p>`),
                    fonts: FONTS,
                }).paragraphs[0]!;

                return isTable(block) ? undefined : block.runs[0]?.shape;
            };

            // White and black are what a shape saying nothing gets.
            expect(shapeOf('')).toMatchObject({ fillHex: '#FFFFFF', strokeHex: '#000000' });
            expect(shapeOf('fillcolor="#ff0000"')?.fillHex).toBe('#FF0000');
            // The short spelling CSS allows, which VML allows too.
            expect(shapeOf('fillcolor="#f00"')?.fillHex).toBe('#FF0000');
            // `f` is how VML says no.
            expect(shapeOf('filled="f"')?.fillHex).toBeUndefined();
            expect(shapeOf('stroked="f"')?.strokeHex).toBeUndefined();
            expect(shapeOf('strokeweight="3pt"')?.strokeWidthPx).toBeCloseTo(4, 6);
        });

        it('draws nothing for a shape that is neither filled nor stroked', () => {
            const block = readWordDocument({
                documentXml: doc(`<w:p><w:r><w:pict><v:rect ${NS_V} `
                    + 'style="width:36pt;height:12pt" filled="f" stroked="f"/>'
                    + '</w:pict></w:r></w:p>'),
                fonts: FONTS,
            }).paragraphs[0]!;
            const runs = isTable(block) ? [] : block.runs;

            expect(runs.every((run) => undefined === run.shape)).toBe(true);
        });

        it('refuses a colour NAME rather than guessing at it', () => {
            // Leaving the shape unfilled beats filling it in a colour nobody
            // asked for.
            const block = readWordDocument({
                documentXml: doc(`<w:p><w:r><w:pict><v:rect ${NS_V} `
                    + 'style="width:36pt;height:12pt" fillcolor="red"/></w:pict></w:r></w:p>'),
                fonts: FONTS,
            }).paragraphs[0]!;
            const runs = isTable(block) ? [] : block.runs;

            expect(runs[0]!.shape?.fillHex).toBeUndefined();
            expect(runs[0]!.shape?.strokeHex).toBe('#000000');
        });

        it('keeps the room but paints nothing for a shape naming no shapetype', () => {
            // A `v:shape` borrows its outline from the `v:shapetype` its
            // `type` names. Naming none leaves no path, and LibreOffice draws
            // nothing — even for the loudest colours a shape can ask for.
            //
            // Measured with a control in one print: a `v:rect`
            // carrying these very attributes came out with its stroke at
            // width 3.000, and this shape beside it made no mark at all.
            // Before that print this engine drew a white box with a black
            // border here, in a document that shows neither.
            const shapeOf = (element: string) => {
                const block = readWordDocument({
                    documentXml: doc(`<w:p><w:r><w:pict>${element}</w:pict></w:r></w:p>`),
                    fonts: FONTS,
                }).paragraphs[0]!;

                return isTable(block) ? undefined : block.runs[0]?.shape;
            };
            const attributes = 'style="width:90pt;height:36pt" fillcolor="#FF0000"'
                + ' strokecolor="#0000FF" strokeweight="3pt"';

            const bare = shapeOf(`<v:shape ${NS_V} ${attributes}/>`);
            const rect = shapeOf(`<v:rect ${NS_V} ${attributes}/>`);

            // The room is kept — that is what the run after it is measured
            // against — and only the paint is dropped.
            expect(bare?.widthPx).toBeCloseTo(120, 6);
            expect(bare?.heightPx).toBeCloseTo(48, 6);
            expect(bare?.fillHex).toBeUndefined();
            expect(bare?.strokeHex).toBeUndefined();

            // The control, in the same shape of test as in the print: same
            // attributes, an element that has geometry, and it paints.
            expect(rect?.fillHex).toBe('#FF0000');
            expect(rect?.strokeHex).toBe('#0000FF');

            // A shape that NAMES a type is untouched: unmeasured, so unchanged.
            expect(shapeOf(`<v:shape ${NS_V} type="#_x0000_t202" ${attributes}/>`)?.fillHex)
                .toBe('#FF0000');
        });

        it('reports a picture the package did not supply', () => {
            const document = readWordDocument({
                documentXml: doc(`<w:p>${picture('width:1in;height:1in')}</w:p>`),
                fonts: FONTS,
            });

            expect(document.diagnostics.map((entry) => entry.detail))
                .toContain('a VML picture referenced "rIdImg", which the package did not supply');
        });
    });


    describe('small caps', () => {
        const runsOf = (rPr: string, text: string) => {
            const block = read(doc(
                `<w:p><w:r><w:rPr><w:sz w:val="20"/>${rPr}</w:rPr>`
                + `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
            )).paragraphs[0]!;

            return isTable(block) ? [] : block.runs;
        };

        it('draws lower case as SMALLER capitals', () => {
            const [only] = runsOf('<w:smallCaps/>', 'abc');
            const [plain] = runsOf('', 'abc');

            expect(only!.text).toBe('ABC');
            expect(only!.sizePx).toBeCloseTo(plain!.sizePx * 0.8, 9);
        });

        it('SPLITS the run where its case changes', () => {
            // One run cannot be two sizes. Measured against LibreOffice, the
            // capital kept its size and the letters after it did not, so a run
            // drawn at one size throughout would be the wrong width either way.
            const runs = runsOf('<w:smallCaps/>', 'Abc');
            const [plain] = runsOf('', 'x');

            expect(runs.map((run) => run.text)).toEqual(['A', 'BC']);
            expect(runs[0]!.sizePx).toBe(plain!.sizePx);
            expect(runs[1]!.sizePx).toBeCloseTo(plain!.sizePx * 0.8, 9);
        });

        it('leaves a run already in capitals alone', () => {
            const runs = runsOf('<w:smallCaps/>', 'ABC');
            const [plain] = runsOf('', 'x');

            expect(runs.map((run) => run.text)).toEqual(['ABC']);
            expect(runs[0]!.sizePx).toBe(plain!.sizePx);
        });

        it('keeps digits and punctuation at full size', () => {
            // Uppercasing them changes nothing, so making them smaller would
            // shrink half a sentence for no reason a reader could see.
            const runs = runsOf('<w:smallCaps/>', 'a1.b');
            const [plain] = runsOf('', 'x');

            expect(runs.map((run) => run.text)).toEqual(['A', '1.', 'B']);
            expect(runs[1]!.sizePx).toBe(plain!.sizePx);
        });

        it('lets w:caps win, which reduces nothing', () => {
            // Both capitalise; only one of them makes anything smaller, and a
            // run that did both would draw ALL CAPS at four fifths.
            const runs = runsOf('<w:caps/><w:smallCaps/>', 'abc');
            const [plain] = runsOf('', 'x');

            expect(runs.map((run) => run.text)).toEqual(['ABC']);
            expect(runs[0]!.sizePx).toBe(plain!.sizePx);
        });
    });


    describe('character spacing', () => {
        const runOf = (rPr: string) => {
            const block = read(doc(
                `<w:p><w:r><w:rPr><w:sz w:val="20"/>${rPr}</w:rPr>`
                + '<w:t xml:space="preserve">ab</w:t></w:r></w:p>',
            )).paragraphs[0]!;

            return isTable(block) ? undefined : block.runs[0];
        };

        it('reads w:spacing as TWENTIETHS of a point', () => {
            // The unit nothing else in the file uses: `w:sz` counts halves and
            // the indents count twentieths of a TWIP.
            expect(runOf('<w:spacing w:val="40"/>')?.letterSpacingPx)
                .toBeCloseTo(2 * 96 / 72, 9);
            expect(runOf('<w:spacing w:val="-20"/>')?.letterSpacingPx)
                .toBeCloseTo(-96 / 72, 9);
        });

        it('leaves a run that asks for none without any', () => {
            // Absent rather than nought, so nothing downstream has to multiply
            // by a zero it was given for every run in the document.
            expect(runOf('')?.letterSpacingPx).toBeUndefined();
            expect(runOf('<w:spacing w:val="0"/>')?.letterSpacingPx).toBeUndefined();
        });
    });


    describe('endnotes', () => {
        const NS_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
        const notesXml = (...ids: number[]): string =>
            `<w:endnotes ${NS_W}>`
            + '<w:endnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r>'
            + '</w:p></w:endnote>'
            + ids.map((id) => `<w:endnote w:id="${id}"><w:p>`
                + '<w:r><w:endnoteRef/></w:r>'
                + `<w:r><w:t xml:space="preserve"> note ${id}</w:t></w:r>`
                + '</w:p></w:endnote>').join('')
            + '</w:endnotes>';

        const mark = (id: number): string => `<w:r><w:endnoteReference w:id="${id}"/></w:r>`;

        const readNoted = (body: string, ...ids: number[]) => readWordDocument({
            documentXml: doc(body),
            fonts: FONTS,
            endnotesXml: notesXml(...ids),
        });

        const textOf = (block: Block | undefined): string =>
            undefined === block || isTable(block)
                ? ''
                : block.runs.map((run) => run.text).join('');

        it('numbers the marks in lower-case ROMAN', () => {
            // Word's default for endnotes, and what LibreOffice printed: an
            // `i` in the text is an endnote where a `1` is a footnote.
            const document = readNoted(
                `<w:p><w:r><w:t>a</w:t></w:r>${mark(4)}</w:p>`
                + `<w:p><w:r><w:t>b</w:t></w:r>${mark(9)}</w:p>`,
                4, 9,
            );

            expect(textOf(document.paragraphs[0])).toBe('ai');
            expect(textOf(document.paragraphs[1])).toBe('bii');
        });

        it('prints the notes at the END of the document', () => {
            const document = readNoted(
                `<w:p><w:r><w:t>a</w:t></w:r>${mark(1)}</w:p>`
                + '<w:p><w:r><w:t>LAST</w:t></w:r></w:p>',
                1,
            );

            expect(document.paragraphs.map(textOf))
                .toEqual(['ai', 'LAST', 'i note 1']);
        });

        it('opens each note with its own number', () => {
            const document = readNoted(
                `<w:p><w:r><w:t>a</w:t></w:r>${mark(7)}${mark(3)}</w:p>`, 3, 7);

            expect(document.paragraphs.slice(1).map(textOf))
                .toEqual(['i note 7', 'ii note 3']);
        });

        it('reads no note the body never refers to', () => {
            const document = readNoted('<w:p><w:r><w:t>a</w:t></w:r></w:p>', 1, 2);

            expect(document.paragraphs.length).toBe(1);
        });

        it('does not raise a mark the DOCUMENT already raised', () => {
            // Word's built-in style raises it and most files carry that style.
            // Applying the default on top would put the mark twice as far up
            // and draw it at four fifths of four fifths.
            const NS = 'xmlns:w="http://schemas.openxmlformats.org/'
                + 'wordprocessingml/2006/main"';
            const styles = `<w:styles ${NS}><w:style w:type="character" w:styleId="ER">`
                + '<w:name w:val="ER"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                + '</w:style></w:styles>';
            const styled = readWordDocument({
                documentXml: doc('<w:p><w:r><w:rPr><w:rStyle w:val="ER"/></w:rPr>'
                    + '<w:endnoteReference w:id="1"/></w:r></w:p>'),
                stylesXml: styles,
                endnotesXml: notesXml(1),
                fonts: FONTS,
            });
            // Given the SAME styles part as the document above, empty of
            // everything but its existence. A package with a styles part and
            // one without fall back to different sizes, and comparing
            // across the two would measure that rather than the raise.
            const plain = readWordDocument({
                documentXml: doc(`<w:p>${mark(1)}</w:p>`),
                stylesXml: `<w:styles ${NS}><w:docDefaults/></w:styles>`,
                endnotesXml: notesXml(1),
                fonts: FONTS,
            });
            const runOf = (document: ReturnType<typeof readWordDocument>) => {
                const block = document.paragraphs[0]!;

                return isTable(block) ? undefined : block.runs[0];
            };

            // The same place either way: the style raised it, or this did.
            expect(runOf(styled)?.baselineShiftPx)
                .toBeCloseTo(runOf(plain)?.baselineShiftPx ?? 0, 9);
            expect(runOf(styled)?.sizePx).toBeCloseTo(runOf(plain)?.sizePx ?? 0, 9);
        });

        it('reports a reference whose note the package did not supply', () => {
            const document = readNoted(`<w:p>${mark(5)}</w:p>`, 1);

            expect(document.diagnostics.map((entry) => entry.detail))
                .toContain('w:endnoteReference "5" has no endnote');
        });
    });


    describe('w:sym', () => {
        const symbol = (font: string, char: string): string =>
            `<w:p><w:r><w:t>A</w:t><w:sym w:font="${font}" w:char="${char}"/>`
            + '<w:t>B</w:t></w:r></w:p>';

        const runsOf = (xml: string) => {
            const block = read(doc(xml)).paragraphs[0]!;

            return isTable(block) ? [] : block.runs;
        };

        it('draws the character a font the engine HAS is asked for', () => {
            // 0x2022 is a bullet, and it is an ordinary code point rather than
            // one of Word's private-use ones.
            const runs = runsOf(symbol('Liberation Serif', '2022'));

            expect(runs.map((run) => run.text)).toEqual(['A', '\u2022', 'B']);
        });

        it('takes the F000 block back out', () => {
            // Word writes a symbol font's own byte with F000 added, which is a
            // convention rather than a Unicode meaning: F041 is the font's
            // 0x41, and looking up F041 would find nothing at all.
            const runs = runsOf(symbol('Liberation Serif', 'F041'));

            expect(runs[1]!.text).toBe('A');
        });

        it('reports a symbol whose font is not here, rather than guessing', () => {
            // The glyph lives only in the font the document names. Drawing its
            // code point out of a substitute prints whatever happens to sit
            // there — an `a` where the document wanted an arrow.
            const document = read(doc(symbol('Wingdings', 'F0E0')));

            expect(document.diagnostics.map((entry) => entry.detail))
                .toContain('a w:sym in "Wingdings" was dropped; the font is not available');
        });

        it('leaves the text around a dropped symbol alone', () => {
            const runs = runsOf(symbol('Wingdings', 'F0E0'));

            expect(runs.map((run) => run.text).join('')).toBe('AB');
        });

        it('passes over a w:char that is not a number', () => {
            // A code that will not parse has no character behind it, and
            // asking for one throws rather than drawing anything.
            const runs = runsOf(symbol('Liberation Serif', 'zz'));

            // One run: nothing was drawn between them, so the text either side
            // was never broken apart.
            expect(runs.map((run) => run.text)).toEqual(['AB']);
        });

        it('draws it in the font the SYMBOL names, not the run own', () => {
            // The glyph is chosen by code, and the code means different things
            // in different fonts — which is the whole reason `w:font` is there.
            const runs = runsOf(symbol('Liberation Mono', '2022'));

            expect(runs[1]!.font).not.toBe(runs[0]!.font);
            expect(runs[2]!.font).toBe(runs[0]!.font);
        });

        it('passes over a w:sym that names no character', () => {
            const runs = runsOf('<w:p><w:r><w:t>A</w:t><w:sym w:font="Symbol"/></w:r></w:p>');

            expect(runs.map((run) => run.text)).toEqual(['A']);
        });
    });


    describe('w:object and w:ruby', () => {
        const NS_V = 'xmlns:v="urn:schemas-microsoft-com:vml"';

        const runsOf = (xml: string, media = true) => {
            const block = readWordDocument({
                documentXml: doc(xml),
                fonts: FONTS,
                ...(media
                    ? { mediaById: { rIdImg: { bytes: new Uint8Array([1]), contentType: 'image/png' } } }
                    : {}),
            }).paragraphs[0]!;

            return isTable(block) ? [] : block.runs;
        };

        it('keeps the picture an embedded OBJECT is stored with', () => {
            // A chart or an equation is an object this cannot run, and Word
            // stores a VML picture of what it looks like beside it. The
            // picture is what a reader sees; dropping it lost the whole thing.
            const runs = runsOf('<w:p><w:r><w:t>before</w:t><w:object>'
                + `<v:shape ${NS_V} style="width:36pt;height:36pt">`
                + '<v:imagedata r:id="rIdImg"/></v:shape></w:object></w:r></w:p>');

            expect(runs.map((run) => run.text)).toEqual(['before', '\uFFFC']);
            expect(runs[1]!.image?.widthPx).toBe(48);
        });

        it('keeps the BASE text of a w:ruby', () => {
            // The base is the sentence; the gloss is set above it. Passing the
            // whole element over took the sentence with it.
            const runs = runsOf('<w:p><w:r><w:ruby>'
                + '<w:rt><w:r><w:t>gloss</w:t></w:r></w:rt>'
                + '<w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase>'
                + '</w:ruby></w:r></w:p>', false);

            expect(runs.map((run) => run.text)).toEqual(['base']);
        });

        it('keeps the GLOSS too, on the run it is set over', () => {
            // Dropped and reported until LibreOffice was found to print it: a
            // second line above the base, inside a line grown by exactly its
            // height.
            const runs = runsOf('<w:p><w:r><w:ruby>'
                + '<w:rubyPr><w:hps w:val="10"/></w:rubyPr>'
                + '<w:rt><w:r><w:rPr><w:sz w:val="10"/></w:rPr><w:t>gloss</w:t></w:r></w:rt>'
                + '<w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase>'
                + '</w:ruby></w:r></w:p>', false);

            expect(runs[0]!.ruby?.text).toBe('gloss');
            // 10 half-points is 5pt, whatever the base run's own size is.
            expect(runs[0]!.ruby?.sizePx).toBeCloseTo(20 / 3, 6);
        });

        it('takes the gloss’s size from w:hps where its own run is silent', () => {
            const runs = runsOf('<w:p><w:r><w:ruby>'
                + '<w:rubyPr><w:hps w:val="10"/></w:rubyPr>'
                + '<w:rt><w:r><w:t>gloss</w:t></w:r></w:rt>'
                + '<w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase>'
                + '</w:ruby></w:r></w:p>', false);

            expect(runs[0]!.ruby?.sizePx).toBeCloseTo(20 / 3, 6);
        });

        it('says so when a gloss covers several runs, which it cannot', () => {
            // The gloss rides on the FIRST run of the base. A base that changes
            // style half way would want it spread across the runs, and nothing
            // measures that — so it is said rather than done quietly.
            const document = readWordDocument({
                documentXml: doc('<w:p><w:r><w:ruby>'
                    + '<w:rt><w:r><w:t>gloss</w:t></w:r></w:rt>'
                    + '<w:rubyBase><w:r><w:t>ba</w:t></w:r><w:r><w:t>se</w:t></w:r></w:rubyBase>'
                    + '</w:ruby></w:r></w:p>'),
                fonts: FONTS,
            });

            expect(document.diagnostics.map((entry) => entry.detail))
                .toContain('a w:ruby gloss over several runs was set over the first of them');
        });

        it('keeps the text around a ruby, in order', () => {
            // TWO runs in the base, because a base of one cannot tell a loop
            // over them from one that reads only the first.
            const runs = runsOf('<w:p><w:r><w:t>A</w:t><w:ruby>'
                + '<w:rt><w:r><w:t>g</w:t></w:r></w:rt>'
                + '<w:rubyBase><w:r><w:t>ba</w:t></w:r><w:r><w:t>se</w:t></w:r>'
                + '</w:rubyBase>'
                + '</w:ruby><w:t>B</w:t></w:r></w:p>', false);

            expect(runs.map((run) => run.text).join('')).toBe('AbaseB');
        });

        it('passes over a ruby with no base at all', () => {
            const runs = runsOf('<w:p><w:r><w:t>A</w:t><w:ruby>'
                + '<w:rt><w:r><w:t>g</w:t></w:r></w:rt></w:ruby></w:r></w:p>', false);

            expect(runs.map((run) => run.text)).toEqual(['A']);
        });
    });


    describe('content controls', () => {
        const sdt = (content: string): string =>
            '<w:sdt><w:sdtPr><w:alias w:val="field"/></w:sdtPr>'
            + `<w:sdtContent>${content}</w:sdtContent></w:sdt>`;

        const textOf = (block: Block | undefined): string =>
            undefined === block || isTable(block)
                ? ''
                : block.runs.map((run) => run.text).join('');

        it('reads the paragraphs a control wraps', () => {
            // A template marks its fillable regions with these, so a reader
            // that passed over them dropped most of the document.
            const document = read(doc(
                paragraph('before') + sdt(paragraph('inside')) + paragraph('after'),
            ));

            expect(document.paragraphs.map(textOf)).toEqual(['before', 'inside', 'after']);
        });

        it('reads a TABLE a control wraps', () => {
            const table = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + '<w:tr><w:tc>' + paragraph('cell') + '</w:tc></w:tr></w:tbl>';
            const document = read(doc(sdt(table)));

            expect(document.paragraphs.filter(isTable).length).toBe(1);
        });

        it('reads a control INSIDE a table cell', () => {
            const table = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + '<w:tr><w:tc>' + sdt(paragraph('held')) + '</w:tc></w:tr></w:tbl>';
            const outer = read(doc(table)).paragraphs.filter(isTable)[0]!;

            expect(textOf(outer.rows[0]!.cells[0]!.paragraphs[0])).toBe('held');
        });

        it('reads ROWS a control wraps, which is what a repeating section is', () => {
            const row = (text: string): string =>
                '<w:tr><w:tc>' + paragraph(text) + '</w:tc></w:tr>';
            const table = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + row('one') + sdt(row('two')) + '</w:tbl>';
            const outer = read(doc(table)).paragraphs.filter(isTable)[0]!;

            expect(outer.rows.length).toBe(2);
            expect(textOf(outer.rows[1]!.cells[0]!.paragraphs[0])).toBe('two');
        });

        it('flattens a control inside a control', () => {
            // A repeating section is a control whose rows are controls.
            const document = read(doc(sdt(sdt(paragraph('deep')))));

            expect(document.paragraphs.map(textOf)).toEqual(['deep']);
        });

        it('passes over a control with no content at all', () => {
            const document = read(doc(
                paragraph('a') + '<w:sdt><w:sdtPr/></w:sdt>' + paragraph('b'),
            ));

            expect(document.paragraphs.map(textOf)).toEqual(['a', 'b']);
        });
    });


    describe('w:pgNumType', () => {
        const sectionWith = (attributes: string): string =>
            `<w:sectPr><w:pgNumType ${attributes}/>`
            + '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

        const firstOf = (attributes: string) => read(
            doc(paragraph('body') + sectionWith(attributes)),
        ).sections[0]?.firstPageNumber;

        it('reads the number a section RESTARTS at', () => {
            // Front matter numbered i, ii, iii and a body beginning again at 1
            // is two sections and one document.
            expect(firstOf('w:start="1"')).toBe(1);
            expect(firstOf('w:start="7"')).toBe(7);
        });

        it('reads nothing from the empty element Word writes anyway', () => {
            // The fixture Word saved for this repository carries a bare
            // `<w:pgNumType/>`: it says nothing, and must not be read as a
            // restart at nought.
            expect(firstOf('')).toBeUndefined();
        });

        it('reads nothing from a start below zero', () => {
            // There is no page before the first. A negative taken as written
            // would number the section backwards from it.
            expect(firstOf('w:start="-3"')).toBeUndefined();
        });

        it('reads nothing from a section that has no pgNumType', () => {
            const plain = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';

            expect(read(doc(paragraph('body') + plain)).sections[0]?.firstPageNumber)
                .toBeUndefined();
        });
    });


    describe('w:pBdr', () => {
        const bordered = (edges: string): string =>
            `<w:p><w:pPr><w:pBdr>${edges}</w:pBdr></w:pPr>`
            + '<w:r><w:t>boxed</w:t></w:r></w:p>';

        const side = (name: string, extra = ''): string =>
            `<w:${name} w:val="single" w:sz="8" w:color="112233" ${extra}/>`;

        const bordersOf = (edges: string) => {
            const block = read(doc(bordered(edges))).paragraphs[0]!;

            return isTable(block) ? undefined : block.style.borders;
        };

        it('reads the four sides a paragraph draws round itself', () => {
            const borders = bordersOf(
                side('top') + side('left') + side('bottom') + side('right'));

            expect(Object.keys(borders ?? {}).sort())
                .toEqual(['bottom', 'left', 'right', 'top']);
            expect(borders?.top?.colorHex).toBe('#112233');
            // `w:sz` counts eighths of a point, so eight of them is one point.
            expect(borders?.top?.widthPx).toBeCloseTo(96 / 72, 9);
        });

        it('reads w:space as POINTS, where the width beside it counts eighths', () => {
            // Measured against LibreOffice: six points of space moved the box
            // six points further out on every side and left the text alone.
            const borders = bordersOf(side('top', 'w:space="6"'));

            expect(borders?.top?.spacePx).toBeCloseTo(6 * 96 / 72, 9);
        });

        it('leaves the space off a border that asks for none', () => {
            expect(bordersOf(side('top'))?.top?.spacePx).toBeUndefined();
            expect(bordersOf(side('top', 'w:space="0"'))?.top?.spacePx).toBeUndefined();
        });

        it('reads w:between, which is the rule a paragraph shares below it', () => {
            expect(bordersOf(side('between'))?.insideH?.colorHex).toBe('#112233');
        });

        it('lets an insideH the element STATES win over w:between', () => {
            // A paragraph border carries `w:between` and a table's carries
            // `w:insideH`; the two land in one place, and what the element
            // said outright is not overwritten by the one folded into it.
            const borders = bordersOf(
                '<w:insideH w:val="single" w:sz="8" w:color="AABBCC"/>' + side('between'));

            expect(borders?.insideH?.colorHex).toBe('#AABBCC');
        });

        it('reads no border from a paragraph that declares none', () => {
            const block = read(doc('<w:p><w:r><w:t>plain</w:t></w:r></w:p>')).paragraphs[0]!;

            expect(isTable(block) ? undefined : block.style.borders).toBeUndefined();
        });

        it('reads nothing from a border switched OFF', () => {
            // `nil` and `none` are how a style's border is cancelled, and are
            // not the same as the element being absent.
            expect(bordersOf('<w:top w:val="nil"/>')).toBeUndefined();
        });
    });

    describe('table width, indent and alignment', () => {
        const tabled = (properties: string): string =>
            `<w:tbl><w:tblPr>${properties}</w:tblPr>`
            + '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>'
            + '<w:tr><w:tc>' + paragraph('a') + '</w:tc>'
            + '<w:tc>' + paragraph('b') + '</w:tc></w:tr></w:tbl>';

        const tableOf = (properties: string) => {
            const block = read(doc(tabled(properties))).paragraphs[0]!;

            return isTable(block) ? block : undefined;
        };

        it('reads w:tblW in FIFTIETHS of a percent, unresolved', () => {
            // Measured: fifty percent of an A4 text column printed at
            // 269.29pt, which is half of 538.58 — so 2500 is half, not
            // twenty-five hundredths of one.
            // 2500 is half, not twenty-five hundredths of one. The fraction is
            // carried rather than turned into a width: the same table is one
            // width on the page and another inside a cell.
            const table = tableOf('<w:tblW w:w="2500" w:type="pct"/>');

            expect(table?.preferredWidthFraction).toBe(0.5);
            expect(table?.preferredWidthPx).toBeUndefined();
        });

        it('reads w:tblW in twips, which needs nothing to resolve against', () => {
            const table = tableOf('<w:tblW w:w="4000" w:type="dxa"/>');

            expect(table?.preferredWidthPx).toBeCloseTo(4000 * 96 / 1440, 9);
            expect(table?.preferredWidthFraction).toBeUndefined();
        });

        it('takes `auto` and a width of nought as no width at all', () => {
            for (const properties of [
                '<w:tblW w:w="4000" w:type="auto"/>',
                '<w:tblW w:w="0" w:type="dxa"/>',
                '',
            ]) {
                const table = tableOf(properties);

                expect(table?.preferredWidthPx).toBeUndefined();
                expect(table?.preferredWidthFraction).toBeUndefined();
            }
        });

        it('defaults an unstated w:type to twips, as the schema does', () => {
            expect(tableOf('<w:tblW w:w="4000"/>')?.preferredWidthPx)
                .toBeCloseTo(4000 * 96 / 1440, 9);
        });

        it('reads the bidirectional spellings of w:jc as the direction they mean', () => {
            expect(tableOf('<w:jc w:val="center"/>')?.alignment).toBe('center');
            expect(tableOf('<w:jc w:val="right"/>')?.alignment).toBe('right');
            expect(tableOf('<w:jc w:val="end"/>')?.alignment).toBe('right');
            expect(tableOf('<w:jc w:val="start"/>')?.alignment).toBe('left');
            expect(tableOf('<w:jc w:val="both"/>')?.alignment).toBeUndefined();
            expect(tableOf('')?.alignment).toBeUndefined();
        });

        it('reads w:tblInd in twips', () => {
            expect(tableOf('<w:tblInd w:w="720" w:type="dxa"/>')?.indentPx)
                .toBeCloseTo(720 * 96 / 1440, 9);
        });

        it('keeps an indent it cannot resolve as NOUGHT rather than dropping it', () => {
            // Measured: LibreOffice made no distance of a percentage indent
            // but still pulled the table left by its cell margin. Absent and
            // present-but-nought are different placements, so they are
            // different values here.
            expect(tableOf('<w:tblInd w:w="1000" w:type="pct"/>')?.indentPx).toBe(0);
            expect(tableOf('')?.indentPx).toBeUndefined();
        });

        it('resolves a percentage COLUMN against the table it sits in', () => {
            // Half of a 6000-twip table is 3000 twips, whatever the grid says.
            const table = tableOf('<w:tblW w:w="6000" w:type="dxa"/>');
            const withPercentages = read(doc(
                '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/></w:tblPr>'
                + '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="7000"/></w:tblGrid>'
                + '<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>'
                + paragraph('a') + '</w:tc>'
                + '<w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>'
                + paragraph('b') + '</w:tc></w:tr></w:tbl>')).paragraphs[0]!;

            expect(table?.columnWidthsPx.map(Math.round))
                .toEqual([4000, 4000].map((twips) => Math.round(twips * 96 / 1440)));
            expect(isTable(withPercentages) ? withPercentages.columnWidthsPx : [])
                .toEqual([3000, 3000].map((twips) => twips * 96 / 1440));
        });

        it('still reports a percentage column it has no table width for', () => {
            const document = read(doc(
                '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + '<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="pct"/></w:tcPr>'
                + paragraph('a') + '</w:tc></w:tr></w:tbl>'));

            expect(document.diagnostics.map((d) => d.detail))
                .toContain('a table column width of type "pct" fell back to the grid');
        });
    });

    describe('w:tcMar', () => {
        const celled = (properties: string) => {
            const table = read(doc(
                '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + `<w:tr><w:tc><w:tcPr>${properties}</w:tcPr>`
                + paragraph('a') + '</w:tc></w:tr></w:tbl>')).paragraphs[0]!;

            return isTable(table) ? table.rows[0]!.cells[0]!.margins : undefined;
        };

        const twips = (value: number): number => value * 96 / 1440;

        it('reads each side a cell states', () => {
            const margins = celled('<w:tcMar><w:left w:w="500" w:type="dxa"/>'
                + '<w:top w:w="400" w:type="dxa"/></w:tcMar>');

            expect(margins?.leftPx).toBeCloseTo(twips(500), 9);
            expect(margins?.topPx).toBeCloseTo(twips(400), 9);
        });

        it('leaves a side the cell did not state UNSET, not nought', () => {
            // Absent means the table's margin still applies; nought would
            // override it with no padding at all. Different placements.
            const margins = celled('<w:tcMar><w:left w:w="500" w:type="dxa"/></w:tcMar>');

            expect(margins?.rightPx).toBeUndefined();
            expect(margins?.bottomPx).toBeUndefined();
        });

        it('reads a `nil` side as nought, which IS an override', () => {
            // The width is 500 and the answer is still nought: `nil` means the
            // cell pads itself by nothing, whatever number sits beside it.
            // Word writes `w:w="0"` there, which would let the rule pass
            // untested — so this states a width the rule has to throw away.
            expect(celled('<w:tcMar><w:left w:w="500" w:type="nil"/></w:tcMar>')?.leftPx)
                .toBe(0);
        });

        it('reads nothing at all from a cell with no w:tcMar', () => {
            expect(celled('')).toBeUndefined();
            expect(celled('<w:tcMar/>')).toBeUndefined();
        });
    });

    describe('w:position', () => {
        const shifted = (properties: string) => read(doc(
            '<w:p><w:r><w:t>a</w:t></w:r>'
            + `<w:r><w:rPr>${properties}</w:rPr><w:t>b</w:t></w:r></w:p>`))
            .paragraphs[0]!;

        const runOf = (properties: string) => {
            const block = shifted(properties);

            return isTable(block) ? undefined : block.runs[1];
        };

        it('reads a raise as a NEGATIVE shift, y counting downward', () => {
            expect(runOf('<w:position w:val="12"/>')?.baselineShiftPx).toBeLessThan(0);
            expect(runOf('<w:position w:val="-12"/>')?.baselineShiftPx).toBeGreaterThan(0);
        });

        it('leaves a run alone at nought, and where the element is absent', () => {
            expect(runOf('<w:position w:val="0"/>')?.baselineShiftPx).toBeUndefined();
            expect(runOf('')?.baselineShiftPx).toBeUndefined();
        });

        it('does not let a stated NOUGHT cancel a superscript', () => {
            // A run asking for no movement is not a run asking to be moved, so
            // it has nothing to win over. Probed: LibreOffice drew `position=0`
            // beside a superscript exactly as it drew the superscript alone —
            // same 3.95pt rise, same shrunken width, the run after it starting
            // in the same place.
            const both = runOf('<w:position w:val="0"/><w:vertAlign w:val="superscript"/>');
            const script = runOf('<w:vertAlign w:val="superscript"/>');

            expect(both?.baselineShiftPx).toBe(script?.baselineShiftPx);
            expect(both?.sizePx).toBe(script?.sizePx);
        });

        it('keeps the run at FULL size, unlike a script', () => {
            const plain = runOf('')!;

            expect(runOf('<w:position w:val="12"/>')?.sizePx).toBe(plain.sizePx);
            expect(runOf('<w:vertAlign w:val="superscript"/>')?.sizePx)
                .toBeLessThan(plain.sizePx);
        });

        it('scales the shift by the FONT, not by the point size', () => {
            // Two sizes of one font raise a run the same distance; the ratio
            // that decides it is the font's line height over its size.
            const small = runOf('<w:position w:val="12"/><w:sz w:val="20"/>');
            const large = runOf('<w:position w:val="12"/><w:sz w:val="40"/>');

            expect(large?.baselineShiftPx).toBeCloseTo(small!.baselineShiftPx!, 9);
        });
    });

    describe('w:pgBorders', () => {
        const sectioned = (properties: string) => read(doc(
            paragraph('a')
            + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${properties}</w:sectPr>`))
            .sections[0]!;

        const box = (offsetFrom = ''): string =>
            `<w:pgBorders ${offsetFrom}>`
            + '<w:top w:val="single" w:sz="24" w:space="24" w:color="112233"/>'
            + '<w:left w:val="single" w:sz="24" w:space="24" w:color="112233"/>'
            + '</w:pgBorders>';

        it('reads the sides, their width in eighths and their space in POINTS', () => {
            const borders = sectioned(box()).pageBorders?.borders;

            expect(borders?.top?.colorHex).toBe('#112233');
            expect(borders?.top?.widthPx).toBeCloseTo(3 * 96 / 72, 9);
            expect(borders?.top?.spacePx).toBeCloseTo(24 * 96 / 72, 9);
        });

        it('defaults w:offsetFrom to `text`, which is 22.6pt from the other', () => {
            expect(sectioned(box()).pageBorders?.offsetFrom).toBe('text');
            expect(sectioned(box('w:offsetFrom="page"')).pageBorders?.offsetFrom).toBe('page');
            expect(sectioned(box('w:offsetFrom="text"')).pageBorders?.offsetFrom).toBe('text');
        });

        it('reads nothing from a section that declares no border', () => {
            expect(sectioned('').pageBorders).toBeUndefined();
            expect(sectioned('<w:pgBorders/>').pageBorders).toBeUndefined();
        });
    });

    describe('w:lnNumType', () => {
        const sectioned = (properties: string) => read(doc(
            paragraph('a')
            + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${properties}</w:sectPr>`))
            .sections[0]!.lineNumbering;

        it('reads the count, the start and the distance', () => {
            const numbering = sectioned(
                '<w:lnNumType w:countBy="3" w:start="5" w:distance="360"/>');

            expect(numbering?.countBy).toBe(3);
            expect(numbering?.start).toBe(5);
            expect(numbering?.distancePx).toBeCloseTo(360 * 96 / 1440, 9);
        });

        it('leaves half a centimetre where no distance is stated', () => {
            // Measured at 14.05pt in from the writing area, against the
            // 18 a quarter of an inch would have given.
            expect(sectioned('<w:lnNumType w:countBy="1"/>')?.distancePx)
                .toBeCloseTo(0.5 / 2.54 * 96, 9);
        });

        it('refuses a countBy of nought, which would divide by zero', () => {
            expect(sectioned('<w:lnNumType w:countBy="0"/>')?.countBy).toBe(1);
            expect(sectioned('<w:lnNumType/>')?.countBy).toBe(1);
        });

        it('defaults the restart to newPage, and reads the other two', () => {
            expect(sectioned('<w:lnNumType w:countBy="1"/>')?.restart).toBe('newPage');
            expect(sectioned('<w:lnNumType w:restart="continuous"/>')?.restart)
                .toBe('continuous');
            expect(sectioned('<w:lnNumType w:restart="newSection"/>')?.restart)
                .toBe('newSection');
        });

        it('reads nothing from a section that asks for no numbering', () => {
            expect(sectioned('')).toBeUndefined();
        });
    });

    describe('w:bdr', () => {
        const boxed = (properties: string) => {
            const block = read(doc(
                `<w:p><w:r><w:rPr>${properties}</w:rPr><w:t>a</w:t></w:r></w:p>`))
                .paragraphs[0]!;

            return isTable(block) ? undefined : block.runs[0]?.border;
        };

        it('reads the ONE border a run draws round all four of its sides', () => {
            // `w:bdr` is a single element with one width, one colour and one
            // space — not four sides like `w:pBdr`.
            const border = boxed(
                '<w:bdr w:val="single" w:sz="24" w:space="2" w:color="112233"/>');

            expect(border?.colorHex).toBe('#112233');
            expect(border?.widthPx).toBeCloseTo(3 * 96 / 72, 9);
            expect(border?.spacePx).toBeCloseTo(2 * 96 / 72, 9);
        });

        it('reads no border from a run that declares none, or switches it off', () => {
            expect(boxed('')).toBeUndefined();
            expect(boxed('<w:bdr w:val="nil"/>')).toBeUndefined();
            expect(boxed('<w:bdr w:val="none"/>')).toBeUndefined();
        });

        it('inherits one from the run’s character STYLE', () => {
            // The reason `readBorderSide` moved into the style sheet: read off
            // the run's own `w:rPr` alone, a border set by a style is lost.
            const styles = `<w:styles
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:style w:type="character" w:styleId="Boxed"><w:rPr>
                <w:bdr w:val="single" w:sz="8" w:space="1" w:color="445566"/>
                </w:rPr></w:style></w:styles>`;
            const block = read(doc(
                '<w:p><w:r><w:rPr><w:rStyle w:val="Boxed"/></w:rPr>'
                + '<w:t>a</w:t></w:r></w:p>'), styles).paragraphs[0]!;

            expect(isTable(block) ? undefined : block.runs[0]?.border?.colorHex)
                .toBe('#445566');
        });

        it('lets the run’s own w:bdr beat the style’s', () => {
            const styles = `<w:styles
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:style w:type="character" w:styleId="Boxed"><w:rPr>
                <w:bdr w:val="single" w:sz="8" w:space="1" w:color="445566"/>
                </w:rPr></w:style></w:styles>`;
            const block = read(doc(
                '<w:p><w:r><w:rPr><w:rStyle w:val="Boxed"/>'
                + '<w:bdr w:val="single" w:sz="8" w:space="1" w:color="778899"/></w:rPr>'
                + '<w:t>a</w:t></w:r></w:p>'), styles).paragraphs[0]!;

            expect(isTable(block) ? undefined : block.runs[0]?.border?.colorHex)
                .toBe('#778899');
        });
    });

    describe('w:textDirection', () => {
        const turned = (properties: string) => {
            const table = read(doc(
                '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + `<w:tr><w:tc><w:tcPr>${properties}</w:tcPr>`
                + paragraph('a') + '</w:tc></w:tr></w:tbl>')).paragraphs[0]!;

            return isTable(table) ? table.rows[0]!.cells[0]!.textDirection : undefined;
        };

        it('reads both quarter turns', () => {
            // Off the PDF's text matrix: `btLr` came out `[0 1 -1 0]` and
            // `tbRl` `[0 -1 1 0]` — anticlockwise and clockwise.
            expect(turned('<w:textDirection w:val="btLr"/>')).toBe('btLr');
            expect(turned('<w:textDirection w:val="tbRl"/>')).toBe('tbRl');
        });

        it('leaves a cell upright where it says nothing, or something else', () => {
            expect(turned('')).toBeUndefined();
            // `lrTb` is upright, and the vertical-East-Asian spellings are not
            // quarter turns of Latin text — none of them is read as one.
            expect(turned('<w:textDirection w:val="lrTb"/>')).toBeUndefined();
            expect(turned('<w:textDirection w:val="tbRlV"/>')).toBeUndefined();
        });
    });

    describe('what it cannot do yet, it says', () => {
        it('no longer reports a table nested inside a cell, which it now reads', () => {
            const nested = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>'
                + '<w:tbl><w:tr><w:tc>' + paragraph('inner') + '</w:tc></w:tr></w:tbl>'
                + paragraph('') + '</w:tc></w:tr></w:tbl>';

            const document = read(doc(nested));
            const outer = document.paragraphs.filter(isTable)[0]!;
            const held = outer.rows[0]!.cells[0]!.paragraphs[0]!;

            expect(document.diagnostics.map((d) => d.detail))
                .not.toContain('a table nested inside a cell was skipped');
            expect(isTable(held)).toBe(true);
        });

        it('no longer reports a vertically merged cell, which it now honours', () => {
            const merged = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>'
                + '<w:tcPr><w:vMerge w:val="restart"/></w:tcPr>' + paragraph('merged')
                + '</w:tc></w:tr></w:tbl>';

            expect(read(doc(merged)).diagnostics.map((d) => d.detail))
                .not.toContain('a vertically merged table cell was laid out as an ordinary one');
        });

        it('reports a CONTINUOUS section break that changes the paper', () => {
            // w:type describes how the section it DEFINES begins, so a
            // continuous break sits on the FOLLOWING section's sectPr. That
            // section shares the previous page, which means the paper it asks
            // for cannot take effect where it sits — folding it in keeps what
            // the document already did, and changing the page a paragraph early
            // would be worse than saying so.
            const first = paragraph('x', '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                + '</w:sectPr></w:pPr>');

            const document = read(doc(first + paragraph('after')
                + '<w:sectPr><w:type w:val="continuous"/>'
                + '<w:pgSz w:orient="landscape" w:w="16838" w:h="11906"/></w:sectPr>'));

            expect(document.diagnostics.map((d) => d.kind)).toContain('unsupported-section-break');
            expect(document.sections.length).toBe(1);
        });

        it('does NOT report one for an ordinary paragraph', () => {
            // A diagnostic that fires on everything is worse than none: it
            // trains the reader to ignore the list. This fired on every
            // paragraph with no w:pPr at all, because optional chaining yields
            // undefined and `null !== undefined` is true.
            const withoutProperties = read(doc('<w:p><w:r><w:t>x</w:t></w:r></w:p>'));
            const withProperties = read(doc(paragraph('x', '<w:pPr><w:keepLines/></w:pPr>')));

            expect(withoutProperties.diagnostics.map((d) => d.kind)).not.toContain('unsupported-section-break');
            expect(withProperties.diagnostics.map((d) => d.kind)).not.toContain('unsupported-section-break');
        });

        it('reports nothing at all for a document it fully understands', () => {
            // The strongest form: a clean document produces an EMPTY list.
            const document = read(doc(paragraph('hello')), `<w:styles
                xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults/></w:styles>`);

            expect(document.diagnostics).toEqual([]);
        });

        it('reports each problem ONCE, however many paragraphs hit it', () => {
            const document = read(doc(paragraph('a', '<w:pPr><w:pStyle w:val="Nope"/></w:pPr>').repeat(50)));

            expect(document.diagnostics.length).toBe(1);
        });

        it('refuses something that is not a WordprocessingML document', () => {
            expect(() => read('<w:document xmlns:w="x"><w:notBody/></w:document>')).toThrow(/no <w:body>/);
        });
    });

    describe('page geometry', () => {
        it('falls back to A4 with one-inch margins when a section declares nothing', () => {
            const document = read(doc(paragraph('x')));

            expect(document.geometry.widthPx).toBe(twipsToPx(11906));
            expect(document.geometry.heightPx).toBe(twipsToPx(16838));
            expect(document.geometry.marginTopPx).toBe(96);
        });

        it('reads each margin separately rather than assuming they match', () => {
            const document = read(doc(
                paragraph('x')
                + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
                + '<w:pgMar w:top="100" w:right="200" w:bottom="300" w:left="400"/></w:sectPr>',
            ));

            expect(document.geometry.marginTopPx).toBe(twipsToPx(100));
            expect(document.geometry.marginRightPx).toBe(twipsToPx(200));
            expect(document.geometry.marginBottomPx).toBe(twipsToPx(300));
            expect(document.geometry.marginLeftPx).toBe(twipsToPx(400));
        });
    });

    describe('lists', () => {
        const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

        const numberingXml = (levels: string): string =>
            `<w:numbering ${W}><w:abstractNum w:abstractNumId="3">${levels}</w:abstractNum>`
            + `<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num></w:numbering>`;

        const bulletLevel = '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25cf"/>'
            + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>';
        const decimalLevel = '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
            + '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>';

        const item = (text: string, ilvl = 0, numId = '4'): string =>
            paragraph(text, `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/>`
                + `<w:numId w:val="${numId}"/></w:numPr></w:pPr>`);

        it('puts the marker in the hanging space and the TEXT at the indent', () => {
            // The bullet is drawn 360 twips left of the indent, and every line
            // of the paragraph — including the first — starts AT the indent.
            // Giving the first line the negative indent instead draws its text
            // on top of the bullet, which is what this used to do.
            const document = readNumbered(doc(item('first')), numberingXml(bulletLevel));
            const block = para(document, 0);

            expect(block.marker?.run.text).toBe('\u25cf');
            expect(block.marker?.offsetPx).toBe(twipsToPx(360));
            expect(block.style.indentLeftPx).toBe(twipsToPx(720));
            expect(block.style.indentFirstLinePx).toBe(0);
        });

        it('sends the first line to the next TAB STOP when the marker outgrows its space', () => {
            // A long marker has to go somewhere, and it does NOT push the text
            // snugly clear of itself — the suffix is a tab, so the text lands
            // on the first stop past the marker's right edge — a `10.`
            // overrunning a 6pt hanging indent put its text 36pt further on
            // rather than the 6.5pt it overran by.
            const wide = '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/>'
                + '<w:lvlText w:val="Section %1 of the appendix —"/>'
                + '<w:pPr><w:ind w:left="720" w:hanging="120"/></w:pPr></w:lvl>';

            const document = readNumbered(doc(item('x')), numberingXml(wide));
            const block = para(document, 0);
            const markerPx = block.marker!.run.font
                .measureAdvance(block.marker!.run.text, block.marker!.run.sizePx).widthPt;

            // The marker starts 120 twips left of the indent and runs past it,
            // so the text goes to the first default stop at or after its end.
            const rightPx = twipsToPx(720) - twipsToPx(120) + markerPx;
            const stopPx = Math.ceil(rightPx / DEFAULT_TAB_PX) * DEFAULT_TAB_PX;

            expect(block.style.indentFirstLinePx).toBe(stopPx - twipsToPx(720));
            // Which is further than the overrun, not equal to it — the whole
            // point of the rule, and what the snug push got wrong.
            expect(block.style.indentFirstLinePx).toBeGreaterThan(markerPx - twipsToPx(120));
        });

        it('counts w:suff="space" as part of the marker\'s width', () => {
            // A space is real width and can be what tips the marker past the
            // hanging space; `nothing` is the same marker without it. Compared
            // against each other rather than against `tab`, which is not a
            // width at all but a jump to a stop.
            const level = (suffix: string): string =>
                '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
                + `<w:suff w:val="${suffix}"/>`
                + '<w:pPr><w:ind w:left="720" w:hanging="20"/></w:pPr></w:lvl>';

            const nothing = para(readNumbered(doc(item('x')), numberingXml(level('nothing'))), 0);
            const space = para(readNumbered(doc(item('x')), numberingXml(level('space'))), 0);
            const spacePx = space.marker!.run.font
                .measureAdvance(' ', space.marker!.run.sizePx).widthPt;

            // To within a twip, which is what the indent is stored in.
            expect(space.style.indentFirstLinePx - nothing.style.indentFirstLinePx)
                .toBeCloseTo(spacePx, 1);
        });

        it('reads the strict spellings of w:jc as well as Word\u2019s own', () => {
            // Word writes left/right; the strict schema writes start/end. A
            // reader that knows only one set silently left-aligns half the
            // documents it is given.
            const alignmentOf = (value: string) => para(
                read(doc(paragraph('t', `<w:pPr><w:jc w:val="${value}"/></w:pPr>`))), 0,
            ).style.alignment;

            expect(alignmentOf('start')).toBe('left');
            expect(alignmentOf('end')).toBe('right');
            expect(alignmentOf('left')).toBe('left');
            expect(alignmentOf('right')).toBe('right');
        });

        it('lets an UNKNOWN w:jc fall through to the style rather than overruling it', () => {
            // Answering 'left' for a value we do not know would quietly undo an
            // alignment the style got right — the paragraph would look like it
            // had asked to be left-aligned when it had asked for nothing we
            // understood.
            const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="Centred">`
                + '<w:pPr><w:jc w:val="center"/></w:pPr></w:style></w:styles>';
            const styled = (jc: string) => para(read(
                doc(paragraph('t', `<w:pPr><w:pStyle w:val="Centred"/><w:jc w:val="${jc}"/></w:pPr>`)),
                styles,
            ), 0).style.alignment;

            expect(styled('right')).toBe('right');
            expect(styled('lowInterlineSpacing')).toBe('center');
        });

        it('takes its list from the STYLE when the paragraph names none', () => {
            // Word's own "List Paragraph" works this way, and a reader that
            // looked only at the paragraph would see no list at all.
            const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="ListParagraph">`
                + '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr>'
                + '</w:style></w:styles>';

            const document = readWordDocument({
                documentXml: doc(paragraph('styled', '<w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr>')),
                numberingXml: numberingXml(bulletLevel),
                stylesXml: styles,
                fonts: FONTS,
            });

            expect(para(document, 0).marker?.run.text).toBe('\u25cf');
            expect(para(document, 0).style.indentLeftPx).toBe(twipsToPx(720));
        });

        it('inherits a list DOWN the basedOn chain', () => {
            // A house style based on List Paragraph is still a list, and the
            // paragraph using it names neither the list nor the style that
            // declared it.
            const styles = `<w:styles ${W}>`
                + '<w:style w:type="paragraph" w:styleId="ListParagraph">'
                + '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr></w:style>'
                + '<w:style w:type="paragraph" w:styleId="HouseList">'
                + '<w:basedOn w:val="ListParagraph"/></w:style>'
                + '</w:styles>';

            const document = readWordDocument({
                documentXml: doc(paragraph('inherited', '<w:pPr><w:pStyle w:val="HouseList"/></w:pPr>')),
                numberingXml: numberingXml(bulletLevel),
                stylesXml: styles,
                fonts: FONTS,
            });

            expect(para(document, 0).marker?.run.text).toBe('●');
        });

        it('lets the paragraph\'s own numId beat the style\'s', () => {
            const styles = `<w:styles ${W}><w:style w:type="paragraph" w:styleId="ListParagraph">`
                + '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr>'
                + '</w:style></w:styles>';

            const document = readWordDocument({
                documentXml: doc(paragraph('escaped',
                    '<w:pPr><w:pStyle w:val="ListParagraph"/>'
                    + '<w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>')),
                numberingXml: numberingXml(bulletLevel),
                stylesXml: styles,
                fonts: FONTS,
            });

            // numId 0 is how a paragraph escapes a style that lists it.
            expect(para(document, 0).marker).toBeUndefined();
        });

        it('counts items in document order', () => {
            const document = readNumbered(
                doc(item('one') + item('two') + item('three')),
                numberingXml(decimalLevel),
            );

            expect([0, 1, 2].map((index) => para(document, index).marker?.run.text))
                .toEqual(['1.', '2.', '3.']);
        });

        it('starts counting again when the document is read again', () => {
            // The counters belong to the READ, not to the definitions. Sharing
            // them would make the second open of a document number differently
            // from the first.
            const xml = doc(item('one') + item('two'));
            const numbering = numberingXml(decimalLevel);

            expect(para(readNumbered(xml, numbering), 0).marker?.run.text).toBe('1.');
            expect(para(readNumbered(xml, numbering), 0).marker?.run.text).toBe('1.');
        });

        it('leaves an ordinary paragraph unmarked', () => {
            const document = readNumbered(doc(paragraph('plain')), numberingXml(bulletLevel));

            expect(para(document, 0).marker).toBeUndefined();
            expect(para(document, 0).style.indentLeftPx).toBe(0);
        });

        it('treats numId 0 as NOT a list', () => {
            // That is how a paragraph escapes a style that puts it in one.
            const document = readNumbered(
                doc(item('escaped', 0, '0')),
                numberingXml(bulletLevel),
            );

            expect(para(document, 0).marker).toBeUndefined();
            // And it is not a BROKEN reference: reporting it would fill the list
            // with noise for something the document said on purpose.
            expect(document.diagnostics.map((d) => d.kind)).not.toContain('unknown-numbering');
        });

        it('reports a numId the numbering part does not define', () => {
            const document = readNumbered(doc(item('orphan', 0, '99')), numberingXml(bulletLevel));

            expect(para(document, 0).marker).toBeUndefined();
            expect(document.diagnostics).toContainEqual({
                kind: 'unknown-numbering',
                detail: 'w:numId "99" is not defined in numbering.xml',
            });
        });

        it('lets the paragraph\'s own indent beat the level\'s', () => {
            // Numbering sits between the style and direct formatting, so a
            // paragraph can still push one bullet further in than its siblings.
            const overridden = paragraph('pushed',
                '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>'
                + '<w:ind w:left="1440"/></w:pPr>');

            const document = readNumbered(doc(overridden), numberingXml(bulletLevel));

            expect(para(document, 0).style.indentLeftPx).toBe(twipsToPx(1440));
            expect(para(document, 0).marker?.run.text).toBe('\u25cf');
        });

        it('marks only the FIRST half of an item split by a page break', () => {
            // The continuation is the same list item; numbering it again would
            // invent an entry that is not in the document.
            const split = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr></w:pPr>'
                + '<w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>';

            const document = readNumbered(doc(split), numberingXml(decimalLevel));

            expect(document.paragraphs.length).toBe(2);
            expect(para(document, 0).marker?.run.text).toBe('1.');
            expect(para(document, 1).marker).toBeUndefined();
        });

        it('does not let the marker change where the text wraps', () => {
            // The marker is drawn in the hanging indent, not in the text: two
            // paragraphs of the same words in the same column must break the
            // same way whether or not one of them is a list.
            const numbering = numberingXml(decimalLevel);
            const text = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda';

            const listed = para(readNumbered(doc(item(text)), numbering), 0);
            const plain = para(readNumbered(doc(paragraph(text)), numbering), 0);

            expect(listed.runs.map((run) => run.text)).toEqual(plain.runs.map((run) => run.text));
        });
    });

        it('carries the line RULE through, not only the height it produced', () => {
            // `exact` and `atLeast` can resolve to the same number and still put
            // the baseline in different places, so the rule has to survive.
            const styled = (rule: string) => para(read(doc(paragraph(
                't', `<w:pPr><w:spacing w:line="480" w:lineRule="${rule}"/></w:pPr>`,
            ))), 0).style;

            expect(styled('exact').lineRule).toBe('exact');
            expect(styled('atLeast').lineRule).toBe('atLeast');
            expect(styled('exact').lineHeightPx).toBe(styled('atLeast').lineHeightPx);
        });

    describe('table borders and shading', () => {
        const tableXml = (tblPr: string, tcPr = ''): string =>
            '<w:tbl><w:tblPr>' + tblPr + '</w:tblPr>'
            + '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
            + `<w:tr><w:tc><w:tcPr>${tcPr}</w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;

        const tableOf = (tblPr: string, tcPr = '') => {
            const block = read(doc(tableXml(tblPr, tcPr))).paragraphs[0]!;
            if (!('rows' in block)) {
                throw new Error('expected a table');
            }

            return block;
        };

        it('reads each side with its own width, colour and style', () => {
            // w:sz is in EIGHTHS of a point: 16 eighths is 2pt, which is 2.667px.
            const table = tableOf(
                '<w:tblBorders>'
                + '<w:top w:val="single" w:sz="16" w:color="FF0000"/>'
                + '<w:insideH w:val="dashed" w:sz="8" w:color="auto"/>'
                + '</w:tblBorders>',
            );

            expect(table.borders?.top).toEqual({
                widthPx: eighthPointsToPx(16),
                colorHex: '#FF0000',
                style: 'solid',
            });
            // `auto` is what Word writes for "the usual colour", and draws black.
            expect(table.borders?.insideH).toEqual({
                widthPx: eighthPointsToPx(8),
                colorHex: '#000000',
                style: 'dashed',
            });
        });

        it('treats nil and none as a border switched OFF', () => {
            const table = tableOf(
                '<w:tblBorders><w:top w:val="nil"/><w:left w:val="none"/></w:tblBorders>',
            );

            expect(table.borders).toBeUndefined();
        });

        it('folds an unrecognised style onto a solid rule rather than losing it', () => {
            // An unknown style is still a LINE. Drawing nothing would drop a
            // border the document asked for.
            const table = tableOf(
                '<w:tblBorders><w:top w:val="threeDEngrave" w:sz="8"/></w:tblBorders>',
            );

            expect(table.borders?.top?.style).toBe('solid');
        });

        it('reads a cell’s own borders and shading', () => {
            const table = tableOf(
                '',
                '<w:tcBorders><w:left w:val="single" w:sz="24"/></w:tcBorders>'
                + '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
            );

            expect(table.rows[0]!.cells[0]!.borders?.left?.widthPx).toBe(eighthPointsToPx(24));
            expect(table.rows[0]!.cells[0]!.shadingFill).toBe('#D9D9D9');
        });

        it('takes the FILL of shading, not its pattern colour', () => {
            // w:color is the foreground of a hatch pattern this engine does not
            // draw; using it would fill the cell with entirely the wrong colour.
            const table = tableOf('', '<w:shd w:val="pct25" w:color="FF0000" w:fill="00FF00"/>');

            expect(table.rows[0]!.cells[0]!.shadingFill).toBe('#00FF00');
        });

        it('ignores shading of auto, which means none', () => {
            const table = tableOf('', '<w:shd w:val="clear" w:fill="auto"/>');

            expect(table.rows[0]!.cells[0]!.shadingFill).toBeUndefined();
        });
    });

    describe('text colour', () => {
        const colourOf = (rPr: string, stylesXml?: string): string | undefined =>
            para(read(doc(`<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>t</w:t></w:r></w:p>`), stylesXml), 0)
                .runs[0]!.colorHex;

        const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
        const GREEN = `<w:styles ${NS}><w:style w:type="character" w:styleId="G">`
            + '<w:name w:val="G"/><w:rPr><w:color w:val="00FF00"/></w:rPr></w:style></w:styles>';

        it('reads w:color as a hex colour', () => {
            expect(colourOf('<w:color w:val="FF0000"/>')).toBe('#FF0000');
        });

        it('leaves a run that names no colour without one', () => {
            // Absent is not black: it means "whatever the renderer draws in",
            // which a page rendered light-on-dark needs to be able to change.
            expect(colourOf('')).toBeUndefined();
        });

        it('takes a colour from a character STYLE', () => {
            expect(colourOf('<w:rStyle w:val="G"/>', GREEN)).toBe('#00FF00');
        });

        it('lets w:val="auto" OVERRIDE the colour a style supplies', () => {
            // Verified against LibreOffice: the same run prints green without
            // the `auto` and black with it. Treating `auto` as absent would let
            // the style's green show through where the document asked for
            // automatic.
            expect(colourOf('<w:rStyle w:val="G"/><w:color w:val="auto"/>', GREEN))
                .toBe('#000000');
        });

        it('leaves a themed colour to the style rather than inventing one', () => {
            // The theme is in theme1.xml, which this reader is never given.
            expect(colourOf('<w:rStyle w:val="G"/><w:color w:themeColor="accent1"/>', GREEN))
                .toBe('#00FF00');
        });
    });

    describe('pictures', () => {
        const EMU_PER_INCH = 914400;
        const BYTES = new Uint8Array([1, 2, 3]);
        const MEDIA = { rIdImg: { bytes: BYTES, contentType: 'image/png' } };

        /**
         * A `w:drawing`, with the namespace prefixes made deliberately ODD.
         *
         * `wp`, `a` and `pic` are conventions, not rules: a producer may bind
         * those namespaces to any prefix. Matching on the prefix would lose
         * every picture in a file like this one.
         */
        const drawing = (cx: number, cy: number, embed = 'rIdImg', wrapper = 'inline'): string =>
            '<w:drawing>'
            + `<zz:${wrapper} xmlns:zz="http://schemas.openxmlformats.org/drawingml/2006/`
            + 'wordprocessingDrawing">'
            + `<zz:extent cx="${cx}" cy="${cy}"/>`
            + '<zz:docPr id="1" name="p" descr="Alt words"/>'
            + '<q:graphic xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/main">'
            + '<q:graphicData><zp:pic xmlns:zp="http://schemas.openxmlformats.org/drawingml/'
            + '2006/picture"><zp:blipFill>'
            + `<q:blip r:embed="${embed}"/>`
            + `</zp:blipFill></zp:pic></q:graphicData></q:graphic></zz:${wrapper}></w:drawing>`;

        const readPicture = (xml: string) => readWordDocument({
            documentXml: doc(`<w:p><w:r>${xml}</w:r></w:p>`),
            mediaById: MEDIA,
            fonts: FONTS,
        });

        it('reads the size the DOCUMENT asks for, whatever the file is', () => {
            const document = readPicture(drawing(EMU_PER_INCH * 2, EMU_PER_INCH));
            const run = para(document, 0).runs[0]!;

            expect(run.image?.widthPx).toBe(emuToPx(EMU_PER_INCH * 2));
            expect(run.image?.heightPx).toBe(emuToPx(EMU_PER_INCH));
            expect(run.image?.content.bytes).toBe(BYTES);
        });

        it('descends by LOCAL name, so an odd prefix still finds the picture', () => {
            // The fixture above binds wp/a/pic to zz/q/zp on purpose.
            expect(para(readPicture(drawing(100, 100)), 0).runs[0]!.image).toBeDefined();
        });

        it('carries the alternative text', () => {
            expect(para(readPicture(drawing(100, 100)), 0).runs[0]!.image?.content.description)
                .toBe('Alt words');
        });

        it('reports a picture the package did not supply rather than drawing nothing', () => {
            // Silently dropping it would leave a hole in the page with no clue
            // in the file about why.
            const document = readPicture(drawing(100, 100, 'rIdMissing'));

            expect(document.paragraphs[0]).toBeDefined();
            expect(document.diagnostics.map((each) => each.kind)).toContain('unsupported-block');
            expect(document.diagnostics.some((each) => each.detail.includes('rIdMissing'))).toBe(true);
        });

        it('hangs an ANCHORED picture on the paragraph, not in the line', () => {
            // A float occupies no width in the line at all: it is anchored to
            // the paragraph and positioned against it, and the lines step aside
            // for it rather than making room within themselves.
            const document = readPicture(drawing(100, 100, 'rIdImg', 'anchor'));
            const paragraph = para(document, 0);

            expect(paragraph.runs.some((run) => undefined !== run.image)).toBe(false);
            expect(paragraph.floats?.length).toBe(1);
            // Through the ternary rather than a cast, so this also fails if a
            // picture ever comes back as a text box: a box has no image at all
            // and `undefined` is not the bytes.
            const float = paragraph.floats![0]!;
            expect(isFloatingBox(float) ? undefined : float.image.content.bytes).toBe(BYTES);
        });

        it('gives a float back its wrap and its origin', () => {
            const document = readPicture(drawing(100, 100, 'rIdImg', 'anchor'));
            const float = para(document, 0).floats![0]!;

            expect(float.wrap).toBe('square');
            expect(float.behindText).toBe(false);
            // The fixture declares no dist*, so text may sit against it.
            expect(float.marginLeftPx).toBe(0);
            // And no positionH: the safest origin is the writing area, which is
            // where the text it displaces already is.
            expect(float.horizontal.relativeTo).toBe('column');
        });

        it('reads wrapNone as a float the text ignores', () => {
            // It sits over or under the words; nothing steps aside for it.
            const xml = drawing(100, 100, 'rIdImg', 'anchor')
                .replace('<zz:docPr', '<zz:wrapNone/><zz:docPr');

            expect(para(readPicture(xml), 0).floats?.[0]!.wrap).toBe('none');
        });

        it('reads the clear space on every side', () => {
            const xml = drawing(100, 100, 'rIdImg', 'anchor')
                .replace('<zz:anchor', '<zz:anchor distT="12700" distB="25400" '
                    + 'distL="38100" distR="50800"');
            const float = para(readPicture(xml), 0).floats![0]!;

            expect([
                float.marginTopPx, float.marginBottomPx,
                float.marginLeftPx, float.marginRightPx,
            ]).toEqual([emuToPx(12700), emuToPx(25400), emuToPx(38100), emuToPx(50800)]);
        });

        it('gives a float to the FIRST piece of a paragraph a break splits', () => {
            // The continuation is the same paragraph; placing its floats again
            // would draw the picture twice, once on each page.
            const document = readWordDocument({
                documentXml: doc(
                    '<w:p><w:r>' + drawing(100, 100, 'rIdImg', 'anchor')
                    + '<w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>',
                ),
                mediaById: MEDIA,
                fonts: FONTS,
            });

            expect(document.paragraphs.length).toBe(2);
            expect(para(document, 0).floats?.length).toBe(1);
            expect(para(document, 1).floats).toBeUndefined();
        });

        it('passes over a VML shape with no size to draw', () => {
            // No style, so no box: there is nothing to draw and nothing to
            // report either, since no picture was asked for.
            const document = readWordDocument({
                documentXml: doc('<w:p><w:r><w:pict><v:shape/></w:pict></w:r></w:p>'),
                fonts: FONTS,
            });
            const block = document.paragraphs[0]!;
            const runs = isTable(block) ? [] : block.runs;

            expect(runs.every((run) => undefined === run.shape)).toBe(true);
        });

        it('keeps the text either side of the picture in its own runs', () => {
            const document = readWordDocument({
                documentXml: doc(
                    '<w:p><w:r><w:t xml:space="preserve">a </w:t>'
                    + drawing(100, 100)
                    + '<w:t xml:space="preserve"> b</w:t></w:r></w:p>',
                ),
                mediaById: MEDIA,
                fonts: FONTS,
            });
            const runs = para(document, 0).runs;

            expect(runs.map((run) => run.text)).toEqual(['a ', OBJECT_REPLACEMENT, ' b']);
            expect(runs[0]!.image).toBeUndefined();
            expect(runs[1]!.image).toBeDefined();
        });

        describe('text boxes', () => {
            /** A VML text box, with `extra` spliced into the shape's tag. */
            const vmlBox = (text: string, extra = '', inset = ''): string =>
                '<w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"'
                + ` style="position:absolute;margin-left:100pt;margin-top:10pt;${extra}`
                + 'width:120pt;height:60pt">'
                + `<v:textbox${inset}><w:txbxContent>${text}</w:txbxContent></v:textbox>`
                + '</v:shape></w:pict>';

            const boxOf = (xml: string) => {
                const float = para(readWordDocument({
                    documentXml: doc(`<w:p><w:r>${xml}</w:r></w:p>`),
                    mediaById: MEDIA,
                    fonts: FONTS,
                }), 0).floats?.[0];

                if (undefined === float || !isFloatingBox(float)) {
                    throw new Error('not a text box');
                }

                return float;
            };

            it('reads the words inside as blocks of their own', () => {
                const box = boxOf(vmlBox('<w:p><w:r><w:t>inside</w:t></w:r></w:p>'));

                // A box holds BLOCKS — it can hold a table — so the
                // paragraphs have to be picked out of them.
                expect(box.blocks
                    .filter((block): block is Paragraph => !isTable(block))
                    .map((block) => block.runs.map((run) => run.text).join('')))
                    .toEqual(['inside']);
                expect([box.widthPx, box.heightPx]).toEqual([twipsToPx(2400), twipsToPx(1200)]);
                // Nothing steps aside for it unless the shape says so:
                // measured, the body text either side of two VML boxes
                // carrying no `w10:wrap` stepped its ordinary 11.50 and never
                // moved.
                expect(box.wrap).toBe('none');
            });

            it('reads a w10:wrap as text stepping aside', () => {
                const box = boxOf(vmlBox('<w:p/>')
                    .replace('</v:shape>', '<w10:wrap type="square"/></v:shape>'));

                expect(box.wrap).toBe('square');
            });

            it('does not carry off the floats of the paragraph it sits IN', () => {
                // `readParagraph` clears the pending floats when it starts and
                // takes them when it ends — and a box is found part-way through
                // the OUTER paragraph's runs, so reading the paragraphs inside
                // it walks that same collector. Without saving and restoring
                // it, the picture anchored before the box vanishes.
                const document = readWordDocument({
                    documentXml: doc('<w:p><w:r>'
                        + drawing(100, 100, 'rIdImg', 'anchor')
                        + vmlBox('<w:p><w:r><w:t>inside</w:t></w:r></w:p>')
                        + '</w:r></w:p>'),
                    mediaById: MEDIA,
                    fonts: FONTS,
                });
                const floats = para(document, 0).floats ?? [];

                expect(floats.map((float) => isFloatingBox(float) ? 'box' : 'picture'))
                    .toEqual(['picture', 'box']);
            });

            it('keeps a TABLE inside a box, beside the paragraphs', () => {
                // It used to be dropped and said so, because a placed float
                // carried lines and rows wanted the renderer's row path as
                // well — which turned out to be one call. Printed, a
                // 2x2 table in a box at 180pt drew its cells at 261.25 and
                // 321.25 with its rules from 255.30 to 376.30, so there was
                // something real to keep.
                const document = readWordDocument({
                    documentXml: doc('<w:p><w:r>'
                        + vmlBox('<w:p><w:r><w:t>above</w:t></w:r></w:p>'
                            + '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p>'
                            + '</w:tc></w:tr></w:tbl>')
                        + '</w:r></w:p>'),
                    fonts: FONTS,
                });

                expect(document.diagnostics.map((entry) => entry.detail))
                    .not.toContain('a table inside a VML text box was dropped');

                const float = para(document, 0).floats?.[0];
                const blocks = undefined !== float && isFloatingBox(float) ? float.blocks : [];

                // The paragraph AND the table, in the order the box states them.
                expect(blocks.length).toBe(2);
                expect(blocks.filter((block) => isTable(block)).length).toBe(1);
            });

            it('reads the inset the file states, and the default when it is silent', () => {
                // 0.1in across and 0.05in down, which is what LibreOffice
                // printed for a box that stated nothing.
                const tenth = emuToPx(EMU_PER_INCH / 10);
                const twentieth = emuToPx(EMU_PER_INCH / 20);

                expect(boxOf(vmlBox('<w:p/>')).inset).toEqual({
                    leftPx: tenth, topPx: twentieth, rightPx: tenth, bottomPx: twentieth,
                });
                // Each side on its own, and an EMPTY one falls back — which is
                // how VML says "this side only".
                expect(boxOf(vmlBox('<w:p/>', '', ' inset="4pt,,8pt,"')).inset).toEqual({
                    leftPx: twipsToPx(80),
                    topPx: twentieth,
                    rightPx: twipsToPx(160),
                    bottomPx: twentieth,
                });
            });

            it('reads each side of a DrawingML inset off its OWN attribute', () => {
                // Four different values, because four equal ones cannot tell a
                // side read off the wrong attribute from one read off the right
                // one — and `tIns` and `bIns` share a default, so the usual
                // fixture hides exactly that swap.
                const box = boxOf(
                    '<w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/'
                    + 'drawingml/2006/wordprocessingDrawing">'
                    + '<wp:extent cx="1524000" cy="762000"/><wp:wrapNone/>'
                    + '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/'
                    + 'wordprocessingShape">'
                    + '<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>'
                    + '<wps:bodyPr lIns="12700" tIns="25400" rIns="38100" bIns="50800"/>'
                    + '</wps:wsp></wp:anchor></w:drawing>',
                );

                expect(box.inset).toEqual({
                    leftPx: emuToPx(12700),
                    topPx: emuToPx(25400),
                    rightPx: emuToPx(38100),
                    bottomPx: emuToPx(50800),
                });
            });

            it('says so when a VML box names an origin that was never measured', () => {
                // Placed against the column and the paragraph regardless, which
                // is where every measured box sat. Obeying an origin nothing
                // was printed for would be inventing one.
                const document = readWordDocument({
                    documentXml: doc('<w:p><w:r>'
                        + vmlBox('<w:p><w:r><w:t>inside</w:t></w:r></w:p>',
                            'mso-position-horizontal-relative:page;')
                        + '</w:r></w:p>'),
                    fonts: FONTS,
                });

                expect(document.diagnostics.map((entry) => entry.detail)).toContain(
                    'a VML text box named an origin of its own; it was placed against the column'
                    + ' and the paragraph instead',
                );
                expect(para(document, 0).floats?.[0]!.horizontal.relativeTo).toBe('column');
            });
        });
    });

    describe('keeping a paragraph with the next one', () => {
        it('reads w:keepNext', () => {
            const document = read(doc(
                '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>kept</w:t></w:r></w:p>'
                + '<w:p><w:r><w:t>loose</w:t></w:r></w:p>',
            ));

            expect(para(document, 0).style.keepWithNext).toBe(true);
            expect(para(document, 1).style.keepWithNext).toBe(false);
        });

        it('keeps only the LAST piece of a paragraph split by a page break', () => {
            // An explicit break inside a paragraph makes two blocks of it. What
            // follows the first of them is the rest of its own text, not the
            // next paragraph — so only the second piece carries the keep.
            const document = read(doc(
                '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>one</w:t></w:r>'
                + '<w:r><w:br w:type="page"/></w:r><w:r><w:t>two</w:t></w:r></w:p>',
            ));

            expect(para(document, 0).style.keepWithNext).toBe(false);
            expect(para(document, 1).style.keepWithNext).toBe(true);
        });
    });

    describe('off the baseline, and in capitals', () => {
        const scripted = (rPr: string) =>
            para(read(doc(`<w:p><w:r><w:rPr><w:sz w:val="20"/>${rPr}</w:rPr>`
                + '<w:t>ab</w:t></w:r></w:p>')), 0).runs[0]!;

        const plain = scripted('');

        it('makes a script run SMALLER, so every measurement already knows', () => {
            // Reduced here rather than at drawing time: widths, line breaking
            // and the line's own height all follow from the size.
            expect(scripted('<w:vertAlign w:val="superscript"/>').sizePx)
                .toBeCloseTo(plain.sizePx * 0.58, 6);
            expect(scripted('<w:vertAlign w:val="subscript"/>').sizePx)
                .toBeCloseTo(plain.sizePx * 0.58, 6);
        });

        it('raises a superscript and drops a subscript', () => {
            // Negative is upwards. LibreOffice raises 3.95pt and drops 0.9 at
            // ten point, which is where these proportions come from.
            expect(scripted('<w:vertAlign w:val="superscript"/>').baselineShiftPx)
                .toBeCloseTo(-plain.sizePx * 0.395, 6);
            expect(scripted('<w:vertAlign w:val="subscript"/>').baselineShiftPx)
                .toBeCloseTo(plain.sizePx * 0.09, 6);
        });

        it('leaves a run on the line when it says baseline', () => {
            const explicit = scripted('<w:vertAlign w:val="baseline"/>');

            expect(explicit.sizePx).toBe(plain.sizePx);
            expect(explicit.baselineShiftPx).toBeUndefined();
        });

        it('lets baseline CANCEL a superscript the style asked for', () => {
            // Absent and `baseline` look the same on a bare run, so only a
            // style to override can tell them apart — and dropping the value
            // rather than recording it would let the style's superscript
            // through.
            const NS = 'xmlns:w="http://schemas.openxmlformats.org/'
                + 'wordprocessingml/2006/main"';
            const styles = `<w:styles ${NS}><w:style w:type="character" w:styleId="Sup">`
                + '<w:name w:val="Sup"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
                + '</w:style></w:styles>';
            const styled = (rPr: string) => para(read(
                doc(`<w:p><w:r><w:rPr><w:sz w:val="20"/>${rPr}</w:rPr><w:t>ab</w:t></w:r></w:p>`),
                styles,
            ), 0).runs[0]!;

            expect(styled('<w:rStyle w:val="Sup"/>').baselineShiftPx).toBeLessThan(0);
            expect(styled('<w:rStyle w:val="Sup"/><w:vertAlign w:val="baseline"/>')
                .baselineShiftPx).toBeUndefined();
        });

        it('shifts nothing for a run that says nothing', () => {
            expect(plain.baselineShiftPx).toBeUndefined();
        });

        it('draws a w:caps run in CAPITALS, whatever the file stores', () => {
            // Applied to the text rather than at drawing time because it
            // changes the width: measured in lower case, a capitalised run
            // breaks its line early and the page count follows.
            const document = read(doc(
                '<w:p><w:r><w:rPr><w:caps/></w:rPr><w:t>shout</w:t></w:r></w:p>',
            ));

            expect(para(document, 0).runs[0]!.text).toBe('SHOUT');
        });

        it('leaves a run without w:caps as the file stores it', () => {
            const document = read(doc('<w:p><w:r><w:t>quiet</w:t></w:r></w:p>'));

            expect(para(document, 0).runs[0]!.text).toBe('quiet');
        });
    });

    describe('underlining and striking through', () => {
        const decorated = (rPr: string) =>
            para(read(doc(`<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>t</w:t></w:r></w:p>`)), 0).runs[0]!;

        it('reads a single underline', () => {
            expect(decorated('<w:u w:val="single"/>').underline).toEqual({ style: 'single' });
        });

        it('folds the seventeen styles onto the rules it can draw', () => {
            // The PRESENCE of a rule is what a reader sees; whether it waves is
            // not, and dropping it would lose the emphasis entirely.
            const styleOf = (value: string) =>
                decorated(`<w:u w:val="${value}"/>`).underline?.style;

            expect(styleOf('double')).toBe('double');
            expect(styleOf('dotted')).toBe('dotted');
            expect(styleOf('dashLong')).toBe('dashed');
            expect(styleOf('wave')).toBe('single');
            expect(styleOf('thick')).toBe('single');
            // Unrecognised still draws one.
            expect(styleOf('squiggleHeavy')).toBe('single');
        });

        it('treats none as underlining switched OFF', () => {
            // Which is how a run cancels what its style asked for.
            expect(decorated('<w:u w:val="none"/>').underline).toBeUndefined();
        });

        it('reads the rule\u2019s own colour when it differs from the text', () => {
            expect(decorated('<w:u w:val="single" w:color="FF0000"/>').underline)
                .toEqual({ style: 'single', colorHex: '#FF0000' });
            expect(decorated('<w:u w:val="single" w:color="auto"/>').underline)
                .toEqual({ style: 'single' });
        });

        it('reads a single and a double strike', () => {
            expect(decorated('<w:strike/>').strike).toBe('single');
            expect(decorated('<w:dstrike/>').strike).toBe('double');
        });

        it('lets a DOUBLE strike win over a single one', () => {
            // A run carrying both is struck twice in Word, not once.
            expect(decorated('<w:strike/><w:dstrike/>').strike).toBe('double');
        });

        it('leaves an undecorated run alone', () => {
            expect(decorated('').underline).toBeUndefined();
            expect(decorated('').strike).toBeUndefined();
        });
    });

    describe('highlighting', () => {
        const behind = (rPr: string): string | undefined =>
            para(read(doc(`<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>t</w:t></w:r></w:p>`)), 0)
                .runs[0]!.highlightHex;

        it('turns a NAMED highlight into a colour', () => {
            // Checked against LibreOffice's own output: yellow prints as
            // `1 1 0 rg` and green as `0 1 0 rg`.
            expect(behind('<w:highlight w:val="yellow"/>')).toBe('#FFFF00');
            expect(behind('<w:highlight w:val="green"/>')).toBe('#00FF00');
            expect(behind('<w:highlight w:val="darkRed"/>')).toBe('#800000');
        });

        it('treats none as no highlight at all', () => {
            expect(behind('<w:highlight w:val="none"/>')).toBeUndefined();
        });

        it('ignores a name the schema does not allow', () => {
            // The seventeen names are the only legal values, so anything else
            // is a producer's mistake rather than a colour to guess at.
            expect(behind('<w:highlight w:val="chartreuse"/>')).toBeUndefined();
        });

        it('reads a shaded run as the same kind of mark', () => {
            // A marker pen and a shaded run both fill the line behind the text.
            expect(behind('<w:shd w:val="clear" w:color="auto" w:fill="00FFFF"/>')).toBe('#00FFFF');
        });

        it('lets the marker pen WIN over shading', () => {
            // Word draws the highlight over the shading, so the shading is
            // never seen; emitting both would paint one over the other and the
            // order would decide what showed.
            expect(behind('<w:highlight w:val="yellow"/><w:shd w:fill="00FFFF"/>'))
                .toBe('#FFFF00');
        });

        it('ignores shading of auto', () => {
            expect(behind('<w:shd w:val="clear" w:fill="auto"/>')).toBeUndefined();
        });
    });

    describe('fields', () => {
        const complex = (instruction: string, cached: string): string =>
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            + `<w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r>`
            + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            + `<w:r><w:t>${cached}</w:t></w:r>`
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

        /** The runs of the first paragraph, as text and field kind. */
        const runsOf = (documentXml: string) =>
            para(read(doc(documentXml)), 0).runs.map((run) => [run.text, run.field]);

        it('needs the SWITCH, not merely the word', () => {
            // `\\*` in a regular expression is "zero or more backslashes", which
            // matches every word in the instruction — so a field whose name or
            // argument happened to read `roman` was numbered by it without any
            // switch at all.
            const document = read(doc(`<w:p>${complex(' SEQ roman ', '9')}</w:p>`));

            expect(para(document, 0).runs[0]!.fieldFormat).toBeUndefined();
        });

        it('reads the switch on a SIMPLE field too', () => {
            // The two forms carry the instruction in different places, and the
            // simple one is what Word writes for a field nobody has edited.
            const document = read(doc(
                '<w:p><w:fldSimple w:instr=" PAGE  \u005c* ROMAN ">'
                + '<w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>',
            ));

            expect(para(document, 0).runs[0]!.fieldFormat).toBe('upperRoman');
        });

        it('forgets the switch once the field has ENDED', () => {
            // A roman page number followed by a plain total: the second is
            // arabic, and carrying the style over would turn a whole footer
            // roman from the first switch onwards.
            const document = read(doc(
                '<w:p>'
                + complex(' PAGE  \u005c* roman ', '9')
                + complex(' NUMPAGES ', '9')
                + '</w:p>',
            ));
            const runs = para(document, 0).runs;

            expect(runs[0]!.fieldFormat).toBe('lowerRoman');
            expect(runs[1]!.fieldFormat).toBeUndefined();
        });

        it('reads the \u005c* switch that says how to write the number', () => {
            // Verified against LibreOffice, which prints iii, III, c and C for
            // these on the third page.
            const styleOf = (switches: string) =>
                para(read(doc(`<w:p>${complex(` PAGE ${switches}`, '9')}</w:p>`)), 0)
                    .runs[0]!.fieldFormat;

            expect(styleOf(' \u005c* roman ')).toBe('lowerRoman');
            expect(styleOf(' \u005c* ROMAN ')).toBe('upperRoman');
            expect(styleOf(' \u005c* alphabetic ')).toBe('lowerLetter');
            expect(styleOf(' \u005c* ALPHABETIC ')).toBe('upperLetter');
            expect(styleOf(' \u005c* Arabic ')).toBe('decimal');
        });

        it('ignores a switch that is not about NUMBERING', () => {
            // `\u005c* MERGEFORMAT` says to keep the field's formatting when it
            // updates. Word writes it on most fields, so treating it as a
            // numbering would mean reading almost every switch wrongly.
            const document = read(doc(
                `<w:p>${complex(' PAGE  \u005c* MERGEFORMAT ', '9')}</w:p>`,
            ));

            expect(para(document, 0).runs[0]!.field).toBe('page');
            expect(para(document, 0).runs[0]!.fieldFormat).toBeUndefined();
        });

        it('finds the numbering switch among several', () => {
            const document = read(doc(
                `<w:p>${complex(' PAGE  \u005c* MERGEFORMAT  \u005c* roman ', '9')}</w:p>`,
            ));

            expect(para(document, 0).runs[0]!.fieldFormat).toBe('lowerRoman');
        });

        it('marks the RESULT of a complex field, and not its instruction', () => {
            // The instruction is not text and must never be drawn; the cached
            // value is text, and is what gets replaced.
            expect(runsOf(`<w:p>${complex(' PAGE ', '99')}</w:p>`))
                .toEqual([['99', 'page']]);
        });

        it('marks the runs inside a simple field', () => {
            expect(runsOf('<w:p><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>77</w:t>'
                + '</w:r></w:fldSimple></w:p>'))
                .toEqual([['77', 'numPages']]);
        });

        it('reads the field NAME past its switches', () => {
            // `PAGE \* MERGEFORMAT` is still PAGE. Matching the whole
            // instruction would leave every real-world field unrecognised,
            // since Word writes a switch onto most of them.
            expect(runsOf(`<w:p>${complex(' PAGE  \\* MERGEFORMAT ', '99')}</w:p>`))
                .toEqual([['99', 'page']]);
        });

        it('keeps the CACHED value of a field it cannot compute', () => {
            // A TOC page number is the best answer available for a field this
            // engine does not evaluate — better than blanking it.
            expect(runsOf(`<w:p>${complex(' TOC \\o "1-3" ', 'iv')}</w:p>`))
                .toEqual([['iv', undefined]]);
        });

        it('leaves the text either side of a field unmarked', () => {
            const xml = '<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r>'
                + complex(' PAGE ', '99')
                + '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>';

            expect(runsOf(xml)).toEqual([['Page ', undefined], ['99', 'page'], [' end', undefined]]);
        });

        it('stops marking runs once the field has ENDED', () => {
            // The state has to be cleared at `end`, or every run to the end of
            // the paragraph becomes a page number.
            const xml = `<w:p>${complex(' PAGE ', '99')}`
                + '<w:r><w:t>after</w:t></w:r></w:p>';

            expect(runsOf(xml)).toEqual([['99', 'page'], ['after', undefined]]);
        });

        it('does not mark a run after a SIMPLE field either', () => {
            const xml = '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple>'
                + '<w:r><w:t>after</w:t></w:r></w:p>';

            expect(runsOf(xml)).toEqual([['9', 'page'], ['after', undefined]]);
        });
    });

    describe('headers and footers', () => {
        const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

        const hdr = (lines: number, word = 'header'): string =>
            `<w:hdr ${W}>` + Array.from({ length: lines }, () => paragraph(word)).join('') + '</w:hdr>';

        /** A4 with one-inch margins and the header sitting half an inch down. */
        const section = (extra = ''): string =>
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
            + ' w:header="720" w:footer="720"/>' + extra + '</w:sectPr>';

        it('reads a header\'s blocks and measures how tall it is', () => {
            const document = readWithFurniture(doc(paragraph('body') + section()), { default: hdr(2) });
            const header = document.headers.get('default');

            expect(header?.blocks.length).toBe(2);
            expect(header?.heightPx).toBeGreaterThan(0);
            // Two lines are twice one.
            const one = readWithFurniture(doc(paragraph('body') + section()), { default: hdr(1) });
            expect(header?.heightPx).toBeCloseTo((one.headers.get('default')?.heightPx ?? 0) * 2, 6);
        });

        it('leaves the writing area alone when the header FITS its margin', () => {
            // Word draws a header in the margin and only pushes the body once it
            // runs out of room there. A header that always shrank the page would
            // repaginate every document that has one.
            const document = readWithFurniture(doc(paragraph('body') + section()), { default: hdr(1) });
            const box = document.contentBox(0, 1);

            expect(box.topPx).toBe(document.geometry.marginTopPx);
            expect(box.bottomPx).toBe(document.geometry.heightPx - document.geometry.marginBottomPx);
        });

        it('pushes the body down when the header OUTGROWS its margin', () => {
            const document = readWithFurniture(doc(paragraph('body') + section()), { default: hdr(8) });
            const box = document.contentBox(0, 1);
            const header = document.headers.get('default');

            expect(box.topPx).toBeGreaterThan(document.geometry.marginTopPx);
            expect(box.topPx).toBeCloseTo((document.geometry.headerDistancePx ?? 0) + (header?.heightPx ?? 0), 6);
        });

        it('raises the bottom when the footer outgrows ITS margin', () => {
            const document = readWithFurniture(
                doc(paragraph('body') + section()),
                undefined,
                { default: hdr(8, 'footer') },
            );
            const box = document.contentBox(0, 1);

            expect(box.bottomPx).toBeLessThan(document.geometry.heightPx - document.geometry.marginBottomPx);
        });

        it('reads the header and footer DISTANCES from w:pgMar', () => {
            const document = readWithFurniture(doc(paragraph('body') + section()));

            expect(document.geometry.headerDistancePx).toBe(twipsToPx(720));
            expect(document.geometry.footerDistancePx).toBe(twipsToPx(720));
        });

        it('measures a table in a header instead of dropping it', () => {
            // A logo beside an address is a table, and a header that ignored it
            // would be half the height the document says — pushing the body up
            // over furniture that is drawn anyway.
            const withTable = `<w:hdr ${W}>` + paragraph('header line')
                + '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
                + '<w:tr><w:tc>' + paragraph('cell') + '</w:tc></w:tr></w:tbl></w:hdr>';

            const document = readWithFurniture(doc(paragraph('body') + section()), { default: withTable });
            const plain = readWithFurniture(doc(paragraph('body') + section()), { default: hdr(1) });
            const oneLine = plain.headers.get('default')?.heightPx ?? 0;

            // The row holds one line, so the header comes to two of them.
            expect(document.headers.get('default')?.heightPx).toBeCloseTo(oneLine * 2, 6);
        });

        it('has no furniture, and a plain writing area, when none is declared', () => {
            const document = read(doc(paragraph('body') + section()));

            expect(document.headers.size).toBe(0);
            expect(document.contentBox(0, 1).topPx).toBe(document.geometry.marginTopPx);
        });

        describe('which variant a page uses', () => {
            it('gives page one its own header only when w:titlePg says so', () => {
                const titled = readWithFurniture(
                    doc(paragraph('b') + section('<w:titlePg/>')),
                    { default: hdr(1), first: hdr(1, 'first') },
                );
                const plain = readWithFurniture(
                    doc(paragraph('b') + section()),
                    { default: hdr(1), first: hdr(1, 'first') },
                );

                expect(titled.variantForPage(0, 1)).toBe('first');
                // Without the flag the first-page header is never shown, however
                // carefully the document defined it.
                expect(plain.variantForPage(0, 1)).toBe('default');
            });

            it('uses even-page headers only when the document SAYS it does', () => {
                // w:evenAndOddHeaders is off by default, and without it Word
                // ignores an even-page header altogether — even one the section
                // took the trouble to define. Using it regardless would give
                // every other page a header the document asked not to show.
                const off = readWordDocument({
                    documentXml: doc(paragraph('b') + section()),
                    fonts: FONTS,
                    headers: { default: hdr(1), even: hdr(1, 'even') },
                });
                const on = readWordDocument({
                    documentXml: doc(paragraph('b') + section()),
                    fonts: FONTS,
                    headers: { default: hdr(1), even: hdr(1, 'even') },
                    settingsXml: `<w:settings ${W}><w:evenAndOddHeaders/></w:settings>`,
                });

                const walk = (document: WordDocument): FurnitureVariant[] =>
                    [0, 1, 2, 3].map((index) => document.variantForPage(index, index + 1));

                expect(walk(off)).toEqual(['default', 'default', 'default', 'default']);
                // Page NUMBERS start at one, so page two is the first even one.
                expect(walk(on)).toEqual(['default', 'even', 'default', 'even']);
            });

            it('falls back to the default when the variant is not defined', () => {
                const document = readWithFurniture(
                    doc(paragraph('b') + section()),
                    { default: hdr(8) },
                );

                // Page two is 'even' and there is no even header, so it uses the
                // default one — and the writing area is the same on both.
                expect(document.contentBox(1, 2).topPx).toBe(document.contentBox(0, 1).topPx);
            });
        });
    });

    describe('sections', () => {
        const A4 = '<w:pgSz w:w="11906" w:h="16838"/>';
        const LANDSCAPE = '<w:pgSz w:orient="landscape" w:w="16838" w:h="11906"/>';

        const HDR = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

        const headerPart = (lines: number): string =>
            `<w:hdr ${HDR}>` + Array.from({ length: lines }, () => paragraph('h')).join('') + '</w:hdr>';

        /** A sectPr naming a header by relationship id. */
        const withHeader = (id: string | null, pgSz: string): string =>
            '<w:sectPr>'
            + (null === id ? '' : `<w:headerReference w:type="default" r:id="${id}"/>`)
            + pgSz
            + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
            + ' w:header="720" w:footer="720"/></w:sectPr>';

        const readSectioned = (documentXml: string, furnitureById: Record<string, string>) =>
            readWordDocument({ documentXml, furnitureById, fonts: FONTS });

        it('gives each section its OWN header', () => {
            // One short, one tall: the sections have different amounts of room,
            // and a document-level header would give them the same.
            const first = paragraph('a', `<w:pPr>${withHeader('rId1', A4)}</w:pPr>`);
            const document = readSectioned(
                doc(first + paragraph('b') + withHeader('rId2', A4)),
                { rId1: headerPart(1), rId2: headerPart(8) },
            );

            expect(document.sections.length).toBe(2);
            expect(document.sections[0]!.headers.get('default')?.blocks.length).toBe(1);
            expect(document.sections[1]!.headers.get('default')?.blocks.length).toBe(8);

            // The short one fits its margin; the tall one pushes the body down.
            expect(document.sections[0]!.contentBox!(0, 1).topPx)
                .toBe(document.sections[0]!.geometry.marginTopPx);
            expect(document.sections[1]!.contentBox!(0, 1).topPx)
                .toBeGreaterThan(document.sections[1]!.geometry.marginTopPx);
        });

        it('measures each section\'s box against that section\'s OWN paper', () => {
            // The footer's distance is measured from the bottom of the page, so
            // a section on different paper has a different bottom. Using the
            // document's geometry for every section puts the writing area of a
            // landscape page where a portrait one's would be.
            const first = paragraph('a', `<w:pPr>${withHeader('rId1', A4)}</w:pPr>`);
            const document = readSectioned(
                doc(first + paragraph('b') + withHeader('rId1', LANDSCAPE)),
                { rId1: headerPart(1) },
            );

            const portrait = document.sections[0]!;
            const landscape = document.sections[1]!;

            expect(landscape.geometry.heightPx).toBeLessThan(portrait.geometry.heightPx);
            expect(landscape.contentBox!(0, 1).bottomPx)
                .toBe(landscape.geometry.heightPx - landscape.geometry.marginBottomPx);
            expect(landscape.contentBox!(0, 1).bottomPx).toBeLessThan(portrait.contentBox!(0, 1).bottomPx);
        });

        it('re-measures an INHERITED header on the new section\'s paper', () => {
            // The same header, wrapped in two different columns. A landscape
            // section is half as wide again, so a header that runs to two lines
            // in portrait fits on one — and its height is what pushes the body
            // down. Carrying the portrait measurement across starts the
            // landscape body a whole line too low.
            const wide = `<w:hdr ${HDR}>` + paragraph(
                Array.from({ length: 40 }, () => 'aaaaaaaaaa').join(' '),
            ) + '</w:hdr>';

            const first = paragraph('a', `<w:pPr>${withHeader('rId1', A4)}</w:pPr>`);
            const document = readSectioned(
                doc(first + paragraph('b') + withHeader(null, LANDSCAPE)),
                { rId1: wide },
            );

            const portrait = document.sections[0]!.headers.get('default')!;
            const landscape = document.sections[1]!.headers.get('default')!;

            // Same blocks, different height.
            expect(landscape.blocks).toBe(portrait.blocks);
            expect(landscape.heightPx).toBeLessThan(portrait.heightPx);
        });

        it('measures furniture in the COLUMN, not across the sheet', () => {
            // Same paper and the same header; only the margins differ. The
            // writing width is what a header wraps in, so the narrow one must
            // come out taller — measured edge to edge the two are identical.
            const long = `<w:hdr ${HDR}>` + paragraph(
                Array.from({ length: 40 }, () => 'aaaaaaaaaa').join(' '),
            ) + '</w:hdr>';

            const margins = (twips: number): string =>
                '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/>' + A4
                + `<w:pgMar w:top="1440" w:right="${twips}" w:bottom="1440"`
                + ` w:left="${twips}" w:header="720" w:footer="720"/></w:sectPr>`;

            const heightAt = (twips: number): number => readSectioned(
                doc(paragraph('a') + margins(twips)), { rId1: long },
            ).sections[0]!.headers.get('default')!.heightPx;

            expect(heightAt(4000)).toBeGreaterThan(heightAt(1440));
        });

        it('re-measures an inherited FOOTER on the new paper too', () => {
            // The same rule as the header, and worth its own test: a footer is
            // positioned from the bottom edge by its height, so a stale height
            // moves it as well as the body.
            const long = `<w:ftr ${HDR}>` + paragraph(
                Array.from({ length: 40 }, () => 'aaaaaaaaaa').join(' '),
            ) + '</w:ftr>';

            const withFooter = (id: string | null, pgSz: string): string =>
                '<w:sectPr>'
                + (null === id ? '' : `<w:footerReference w:type="default" r:id="${id}"/>`)
                + pgSz
                + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
                + ' w:header="720" w:footer="720"/></w:sectPr>';

            const first = paragraph('a', `<w:pPr>${withFooter('rId1', A4)}</w:pPr>`);
            const document = readSectioned(
                doc(first + paragraph('b') + withFooter(null, LANDSCAPE)),
                { rId1: long },
            );

            const portrait = document.sections[0]!.footers.get('default')!;
            const landscape = document.sections[1]!.footers.get('default')!;

            expect(landscape.blocks).toBe(portrait.blocks);
            expect(landscape.heightPx).toBeLessThan(portrait.heightPx);
        });

        it('INHERITS a header the next section does not declare', () => {
            // Only the first section of a document has to declare everything it
            // uses; a later one that says nothing keeps what came before.
            const first = paragraph('a', `<w:pPr>${withHeader('rId1', A4)}</w:pPr>`);
            const document = readSectioned(
                doc(first + paragraph('b') + withHeader(null, A4)),
                { rId1: headerPart(3) },
            );

            expect(document.sections[1]!.headers.get('default')?.blocks.length).toBe(3);
        });

        it('reports the FIRST section as the document\'s own header', () => {
            // What a caller that knows nothing about sections is looking at.
            const first = paragraph('a', `<w:pPr>${withHeader('rId1', A4)}</w:pPr>`);
            const document = readSectioned(
                doc(first + paragraph('b') + withHeader('rId2', A4)),
                { rId1: headerPart(1), rId2: headerPart(8) },
            );

            expect(document.headers.get('default')?.blocks.length).toBe(1);
        });

        it('ignores a reference to a part that is not there', () => {
            const first = paragraph('a', `<w:pPr>${withHeader('rId9', A4)}</w:pPr>`);
            const document = readSectioned(doc(first + paragraph('b') + withHeader(null, A4)), {});

            expect(document.sections[0]!.headers.size).toBe(0);
            expect(document.sections[0]!.contentBox!(0, 1).topPx)
                .toBe(document.sections[0]!.geometry.marginTopPx);
        });

        describe('which variant a page of a LATER section uses', () => {
            // Section two opens on document page four: its own first page, and
            // an even page of the document. The two indices disagree there, and
            // every rule below was read off LibreOffice printing exactly that.
            const EVEN_AND_ODD =
                `<w:settings ${HDR}><w:evenAndOddHeaders/></w:settings>`;

            const twoSections = (extra: string): string => {
                const properties = (id: string) =>
                    '<w:sectPr>'
                    + `<w:headerReference w:type="default" r:id="${id}Def"/>`
                    + `<w:headerReference w:type="even" r:id="${id}Even"/>`
                    + `<w:headerReference w:type="first" r:id="${id}First"/>`
                    + A4
                    + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
                    + ' w:header="720" w:footer="720"/>' + extra + '</w:sectPr>';

                return doc(
                    paragraph('a', `<w:pPr>${properties('rA')}</w:pPr>`)
                    + paragraph('b') + properties('rB'),
                );
            };

            const parts = Object.fromEntries(
                ['rADef', 'rAEven', 'rAFirst', 'rBDef', 'rBEven', 'rBFirst']
                    .map((id) => [id, headerPart(1)]),
            );

            const secondSection = (extra: string): WordDocument['sections'][number] =>
                readWordDocument({
                    documentXml: twoSections(extra),
                    furnitureById: parts,
                    settingsXml: EVEN_AND_ODD,
                    fonts: FONTS,
                }).sections[1]!;

            it('counts even and odd by the PRINTED page, not by the section', () => {
                // The bug this replaced: page four read as the section's page
                // one, which is odd, so it drew the default header. LibreOffice
                // draws the even one.
                expect(secondSection('').variantForPage(0, 4)).toBe('even');
                expect(secondSection('').variantForPage(1, 5)).toBe('default');
            });

            it('lets w:titlePg BEAT the parity on that same page', () => {
                // Page four is even AND the section's first. LibreOffice draws
                // the first-page header there, so titlePg wins.
                expect(secondSection('<w:titlePg/>').variantForPage(0, 4)).toBe('first');
                expect(secondSection('<w:titlePg/>').variantForPage(1, 5)).toBe('default');
            });

            it('hands layoutSections the blocks of the variant it chose', () => {
                // Reading the variant and DRAWING it are different things: the
                // section has to hand the layout the blocks, or the header is
                // measured and never appears.
                const section = secondSection('<w:titlePg/>');

                expect(section.headerFor?.(0, 4))
                    .toBe(section.headers.get('first')?.blocks);
                expect(section.headerFor?.(1, 5))
                    .toBe(section.headers.get('default')?.blocks);
            });

            it('falls back to the default variant for the BLOCKS as well', () => {
                // Same fallback as the writing area uses. If the two differed,
                // the body would be pushed down by one header's height and a
                // different header would be drawn in the gap.
                const document = readWordDocument({
                    documentXml: twoSections(''),
                    furnitureById: { rADef: headerPart(1), rBDef: headerPart(1) },
                    settingsXml: EVEN_AND_ODD,
                    fonts: FONTS,
                });
                const section = document.sections[1]!;

                expect(section.variantForPage(0, 4)).toBe('even');
                expect(section.headerFor?.(0, 4)).toBe(section.headers.get('default')?.blocks);
            });
        });

        it('gives a document with no breaks exactly one section', () => {
            const document = read(doc(paragraph('a') + paragraph('b') + `<w:sectPr>${A4}</w:sectPr>`));

            expect(document.sections.length).toBe(1);
            expect(document.sections[0]!.blocks.length).toBe(2);
            expect(document.sections[0]!.geometry.widthPx).toBe(twipsToPx(11906));
        });

        it('splits at a mid-document break and gives each section its own paper', () => {
            // A landscape page for a wide table, inside a portrait report.
            const portraitPart = paragraph('portrait', `<w:pPr><w:sectPr>${A4}</w:sectPr></w:pPr>`);
            const document = read(doc(
                portraitPart + paragraph('landscape') + `<w:sectPr>${LANDSCAPE}</w:sectPr>`,
            ));

            expect(document.sections.length).toBe(2);
            expect(document.sections[0]!.geometry.widthPx).toBe(twipsToPx(11906));
            expect(document.sections[1]!.geometry.widthPx).toBe(twipsToPx(16838));
            expect(document.sections[1]!.geometry.widthPx)
                .toBeGreaterThan(document.sections[1]!.geometry.heightPx);
        });

        it('keeps the paragraph that CARRIES the break in the section it ends', () => {
            // The properties describe what came before them, not what follows.
            const document = read(doc(
                paragraph('one') + paragraph('two', `<w:pPr><w:sectPr>${A4}</w:sectPr></w:pPr>`)
                + paragraph('three') + `<w:sectPr>${LANDSCAPE}</w:sectPr>`,
            ));

            expect(document.sections[0]!.blocks.length).toBe(2);
            expect(document.sections[1]!.blocks.length).toBe(1);
        });

        it('numbers the sections\' blocks continuously', () => {
            const document = read(doc(
                paragraph('one') + paragraph('two', `<w:pPr><w:sectPr>${A4}</w:sectPr></w:pPr>`)
                + paragraph('three') + `<w:sectPr>${LANDSCAPE}</w:sectPr>`,
            ));

            expect(document.sections.map((section) => section.firstBlockIndex)).toEqual([0, 2]);
            // And the flat list is still every block in reading order.
            expect(document.paragraphs.length).toBe(3);
        });

        it('reads which page a section has to START on', () => {
            // How a chapter always opens on a right-hand page. Anything the
            // engine does not model — nextColumn, or nothing at all — simply
            // starts a page, which is what a section break does.
            const odd = paragraph('a', `<w:pPr><w:sectPr><w:type w:val="oddPage"/>${A4}</w:sectPr></w:pPr>`);
            const even = paragraph('b', `<w:pPr><w:sectPr><w:type w:val="evenPage"/>${A4}</w:sectPr></w:pPr>`);
            const column = paragraph('c', `<w:pPr><w:sectPr><w:type w:val="nextColumn"/>${A4}</w:sectPr></w:pPr>`);

            const document = read(doc(odd + even + column + paragraph('d') + `<w:sectPr>${A4}</w:sectPr>`));

            expect(document.sections.map((section) => section.startsOn))
                .toEqual(['oddPage', 'evenPage', 'nextPage', 'nextPage']);
        });

        it('folds a continuous break into the section before it', () => {
            // The flag belongs to the section that BEGINS continuously — here
            // the last one, which shares the page with what came before it.
            const first = paragraph('x', `<w:pPr><w:sectPr>${A4}</w:sectPr></w:pPr>`);
            const document = read(doc(
                first + paragraph('after')
                + `<w:sectPr><w:type w:val="continuous"/>${A4}</w:sectPr>`,
            ));

            // No page break, no new section, and nothing to report because the
            // paper did not change.
            expect(document.sections.length).toBe(1);
            expect(document.sections[0]!.blocks.length).toBe(2);
            expect(document.diagnostics.map((d) => d.kind)).not.toContain('unsupported-section-break');
        });

        it('reports a continuous break that changes only the page HEIGHT', () => {
            // Same width and the same margins: only the paper's length differs,
            // which a comparison on width alone would miss entirely.
            const first = paragraph('x', '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                + '</w:sectPr></w:pPr>');

            const document = read(doc(first + paragraph('after')
                + '<w:sectPr><w:type w:val="continuous"/>'
                + '<w:pgSz w:w="11906" w:h="20000"/></w:sectPr>'));

            expect(document.diagnostics.map((d) => d.kind)).toContain('unsupported-section-break');
        });

        it('reports a continuous break that changes only the MARGINS', () => {
            // Same paper, different writing area — which a comparison on page
            // size alone would call identical and silently drop.
            const first = paragraph('x', '<w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
                + '</w:sectPr></w:pPr>');

            const document = read(doc(first + paragraph('after')
                + '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>'
                + '<w:pgMar w:top="2880" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'));

            expect(document.diagnostics.map((d) => d.kind)).toContain('unsupported-section-break');
        });

        it('still starts a NEW section when the break is not continuous', () => {
            // The control for the fold above: the same document without the
            // flag is two sections and two pages.
            const first = paragraph('x', `<w:pPr><w:sectPr>${A4}</w:sectPr></w:pPr>`);
            const document = read(doc(first + paragraph('after') + `<w:sectPr>${A4}</w:sectPr>`));

            expect(document.sections.length).toBe(2);
        });
    });

    describe('tab stops', () => {
        const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

        const tabbed = (pPr: string): ReturnType<typeof read> =>
            read(doc(`<w:p>${pPr}<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>`));

        it('gives a paragraph Word\'s default half-inch stop', () => {
            expect(para(tabbed(''), 0).style.defaultTabPx).toBe(twipsToPx(720));
        });

        it('reads a stop\u2019s LEADER', () => {
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="right" w:pos="1440" w:leader="dot"/>'
                + '<w:tab w:val="right" w:pos="2880" w:leader="underscore"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops?.map((stop) => stop.leader))
                .toEqual(['dot', 'underscore']);
        });

        it('leaves a stop with none, or an unknown one, unfilled', () => {
            // A row of the wrong character is worse than a plain tab.
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="right" w:pos="1440" w:leader="none"/>'
                + '<w:tab w:val="right" w:pos="2880" w:leader="squiggle"/>'
                + '<w:tab w:val="right" w:pos="4320"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops?.map((stop) => stop.leader))
                .toEqual([undefined, undefined, undefined]);
        });

        it('reads the ALIGNMENT of a stop, not only its column', () => {
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="center" w:pos="1440"/>'
                + '<w:tab w:val="right" w:pos="2880"/>'
                + '<w:tab w:val="decimal" w:pos="4320"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops?.map((stop) => stop.align))
                .toEqual(['center', 'right', 'decimal']);
        });

        it('drops a BAR stop, which draws a rule rather than being a stop', () => {
            // A tab never advances to one, so keeping it would stop a column
            // short at a vertical line this engine cannot even draw.
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="bar" w:pos="1440"/><w:tab w:val="left" w:pos="2880"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops)
                .toEqual([{ positionPx: twipsToPx(2880), align: 'left' }]);
        });

        it('reads a numbering stop and an unknown one as LEFT', () => {
            // A stop in the right column aligned the usual way is far closer to
            // the document than no stop at all.
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="num" w:pos="1440"/><w:tab w:val="wobble" w:pos="2880"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops?.map((stop) => stop.align))
                .toEqual(['left', 'left']);
        });

        it('takes the decimal separator from settings.xml', () => {
            const document = readWordDocument({
                documentXml: doc(paragraph('t')),
                settingsXml: `<w:settings ${W}><w:decimalSymbol w:val=","/></w:settings>`,
                fonts: FONTS,
            });

            expect(para(document, 0).style.decimalSymbol).toBe(',');
        });

        it('reads w:tabs, in ascending order', () => {
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="left" w:pos="2880"/><w:tab w:val="left" w:pos="1440"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops).toEqual([
                { positionPx: twipsToPx(1440), align: 'left' },
                { positionPx: twipsToPx(2880), align: 'left' },
            ]);
        });

        it('drops a CLEAR stop, which removes a column rather than adding one', () => {
            const document = tabbed('<w:pPr><w:tabs>'
                + '<w:tab w:val="clear" w:pos="1440"/><w:tab w:val="left" w:pos="2880"/>'
                + '</w:tabs></w:pPr>');

            expect(para(document, 0).style.tabStops)
                .toEqual([{ positionPx: twipsToPx(2880), align: 'left' }]);
        });

        it('takes the default stop from settings.xml when it says one', () => {
            const document = readWordDocument({
                documentXml: doc('<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>'),
                settingsXml: `<w:settings ${W}><w:defaultTabStop w:val="1440"/></w:settings>`,
                fonts: FONTS,
            });

            expect(para(document, 0).style.defaultTabPx).toBe(twipsToPx(1440));
        });

        it('measures a tabbed line by the STOP, not by the tab glyph', () => {
            // End to end: the same two characters, with and without the tab
            // between them, and the difference is a column rather than a glyph.
            const withTab = para(tabbed(''), 0);
            const [page] = layoutPages([withTab], {
                widthPx: 1000, heightPx: 1000,
                marginTopPx: 0, marginRightPx: 0, marginBottomPx: 0, marginLeftPx: 0,
            });

            // The 'b' starts at the 48px stop, so the line is wider than two
            // characters could ever be.
            expect(page!.lines[0]!.line.widthPx).toBeGreaterThan(48);
        });
    });
});
