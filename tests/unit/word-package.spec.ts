import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { OpcPackage, webStreamsCodec, type ZipCodec } from '../../src/opc/opc-package.js';
import { FontCatalogue, type FontManifest } from '../../src/word/font-catalogue.js';
import { openWordFile } from '../../src/word/word-package.js';
import { layoutPages, layoutSections } from '../../src/layout/page-layout.js';
import { XmlDocument } from '../../src/ooxml/xml.js';
import { renderPage, type DrawOp, type ImageOp, type LineOp, type RectOp, type TextOp }
    from '../../src/render/display-list.js';
import { renderPageToSvg } from '../../src/render/svg.js';
import {
    isTable,
    type Block,
    type Page,
    type Paragraph,
    type PlacedFurniture,
    type PlacedParagraphBorder,
    verticalSpans,
} from '../../src/layout/page-layout.js';
import type { LinePiece } from '../../src/layout/line-breaker.js';
import { eighthPointsToPx, twipsToPx } from '../../src/ooxml/units.js';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');
const FONT_DIR = join(ROOT, 'assets/fonts');
const FIXTURES = join(ROOT, 'tests/fixtures/docx');

const MANIFEST = JSON.parse(readFileSync(join(FONT_DIR, 'fonts.manifest.json'), 'utf8')) as FontManifest;
const FONTS = FontCatalogue.load(MANIFEST, (file) => new Uint8Array(readFileSync(join(FONT_DIR, file))));

/** The opened document's block at `index`, as a paragraph. */
function paragraphOf(opened: { document: { paragraphs: readonly Block[] } }, index: number): Paragraph {
    const block = opened.document.paragraphs[index];
    if (undefined === block || isTable(block)) {
        throw new Error(`block ${index} is not a paragraph`);
    }

    return block;
}

function file(name: string): Uint8Array {
    return new Uint8Array(readFileSync(join(FIXTURES, name)));
}

/** The text of a placed header or footer, or null when the page has none. */
function furnitureText(placed: PlacedFurniture | undefined): string | null {
    if (undefined === placed) {
        return null;
    }

    return placed.lines
        .map((line) => line.line.pieces.map((piece) => piece.text).join(''))
        .join(' ');
}

describe('openWordFile', () => {
    it('opens a .docx FILE and paginates it, bytes in', async () => {
        // The whole chain in one call: zip, relationships, XML, styles, fonts,
        // measurement, line breaking, pagination.
        const opened = await openWordFile(file('lease-landscape.docx'), FONTS);
        const pages = layoutPages(opened.document.paragraphs, opened.document.geometry);

        expect(pages.length).toBe(3);
        expect(opened.document.geometry.widthPx).toBe(twipsToPx(16838)); // landscape
    });

    it('follows the relationship to the NUMBERING part as well', async () => {
        // The document's bullet lists are defined in a part the main one points
        // at; without it every bullet is an unmarked, unindented paragraph.
        const opened = await openWordFile(file('word-authored.docx'), FONTS);

        expect(opened.numberingPart).toBe('word/numbering.xml');

        const marked = opened.document.paragraphs.filter((block) => !isTable(block) && undefined !== block.marker);
        expect(marked.length).toBeGreaterThan(0);
        expect(paragraphOf(opened, opened.document.paragraphs.indexOf(marked[0]!)).marker?.run.text)
            .toBe('\u25cf');
    });

    it('loads the header and footer a section references', async () => {
        // Found through the reference's r:id, not by relationship type: a
        // document's default, first-page and even-page headers are all of type
        // "header", and only the reference says which is which.
        const opened = await openWordFile(file('with-header.docx'), FONTS);

        expect(opened.document.headers.get('default')?.blocks.length).toBe(2);
        expect(opened.document.footers.get('default')?.blocks.length).toBe(1);
    });

    it('reads each reference\'s w:type, not just the first header it finds', async () => {
        // Default, first-page and even-page headers are ALL of relationship type
        // "header". Only the reference's w:type tells them apart, so a reader
        // that ignored it would give every page the same one.
        const opc = OpcPackage.open(file('with-header.docx'));
        const xml = await opc.text('word/document.xml');

        expect(xml).toContain('<w:headerReference w:type="default" r:id="rId7"/>');
        opc.replace('word/document.xml', xml.replace(
            '<w:headerReference w:type="default" r:id="rId7"/>',
            '<w:headerReference w:type="default" r:id="rId7"/>'
            + '<w:headerReference w:type="first" r:id="rId7"/><w:titlePg/>',
        ));

        const opened = await openWordFile(await opc.save(), FONTS);

        expect([...opened.document.headers.keys()].sort()).toEqual(['default', 'first']);
        expect(opened.document.variantForPage(0, 1)).toBe('first');
    });

    it('paginates by the writing area the furniture leaves', async () => {
        // Two documents with the same body: one header fits its margin and one
        // does not. LibreOffice makes them three and four pages, and so must
        // this — ignoring the header gives three for both.
        const short = await openWordFile(file('with-header.docx'), FONTS);
        const tall = await openWordFile(file('tall-header.docx'), FONTS);

        const paginate = (opened: Awaited<ReturnType<typeof openWordFile>>): number =>
            layoutPages(opened.document.paragraphs, opened.document.geometry, {
                contentBox: opened.document.contentBox,
            }).length;

        expect(paginate(short)).toBe(3);
        expect(paginate(tall)).toBe(4);
        // The tall header is what makes the difference: the bodies are the same.
        expect(layoutPages(tall.document.paragraphs, tall.document.geometry).length).toBe(3);
    });

    it('lays a multi-section document out on its own papers', async () => {
        // A portrait report with a landscape page in the middle: one document,
        // two page sizes. LibreOffice makes it three pages and draws both box
        // shapes; ignoring the sections gives two pages and one shape, which is
        // the wrong answer to a question the document answered itself.
        const opened = await openWordFile(file('sections.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(opened.document.sections.length).toBe(3);
        expect(pages.length).toBe(3);

        const shape = (index: number): string =>
            `${Math.round(pages[index]!.geometry.widthPx)}x${Math.round(pages[index]!.geometry.heightPx)}`;

        expect([shape(0), shape(1), shape(2)]).toEqual(['794x1123', '1123x794', '794x1123']);
        expect(pages[1]!.geometry.widthPx).toBeGreaterThan(pages[1]!.geometry.heightPx);

        // The control: one geometry for the document is a different answer.
        expect(layoutPages(opened.document.paragraphs, opened.document.geometry).length).toBe(2);
    });

    it('reads a w:orient that CONTRADICTS the page size the way LibreOffice does', async () => {
        // PHPWord writes orient="portrait" beside landscape dimensions. The
        // dimensions win — w:w and w:h are the page, and w:orient is a note
        // about them. Honouring the flag would rotate the page back.
        const opened = await openWordFile(file('sections.docx'), FONTS);
        const landscape = opened.document.sections[1]!;

        expect(landscape.geometry.widthPx).toBeGreaterThan(landscape.geometry.heightPx);
    });

    it('reads only the parts it needs, not every part in the package', async () => {
        // Header and footer parts are collected by relationship id, and only
        // those: collecting every relationship would inflate the comments, the
        // font table and each image as text on open.
        let inflated = 0;
        const counting: ZipCodec = {
            inflateRaw: (bytes) => {
                inflated++;

                return webStreamsCodec.inflateRaw(bytes);
            },
            deflateRaw: (bytes) => webStreamsCodec.deflateRaw(bytes),
        };

        const bytes = file('word-authored.docx');
        await openWordFile(bytes, FONTS, counting);

        // The two relationship parts, the document, its styles, its numbering,
        // its settings, its footnotes and its endnotes. The package holds
        // eighteen.
        expect(inflated).toBe(8);
        expect(OpcPackage.open(bytes).names.length).toBe(18);
    });

    it('inserts the blank page an oddPage break needs, as LibreOffice does', async () => {
        // Chapter two must open on a right-hand page. Chapter one ends on page
        // one, so page two is blank and chapter two starts on page three —
        // three pages for two chapters that would otherwise take two.
        const opened = await openWordFile(file('odd-page-break.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(opened.document.sections.map((section) => section.startsOn))
            .toEqual(['nextPage', 'oddPage']);

        expect(pages.length).toBe(3);
        expect(pages[1]!.lines).toEqual([]);
        expect(pages[1]!.rows).toEqual([]);
        expect(pages[2]!.sectionIndex).toBe(1);
    });

    it('gives each section the header IT references', async () => {
        // Two sections, a one-line header and an eight-line one. The second's
        // header outgrows its margin and its pages hold less; the first's does
        // not. LibreOffice makes the document two pages and so does this.
        const opened = await openWordFile(file('section-headers.docx'), FONTS);
        const [one, two] = opened.document.sections;

        expect(one!.headers.get('default')?.blocks.length).toBe(1);
        expect(two!.headers.get('default')?.blocks.length).toBe(8);

        expect(one!.contentBox!(0, 1).topPx).toBe(one!.geometry.marginTopPx);
        expect(two!.contentBox!(0, 1).topPx).toBeGreaterThan(two!.geometry.marginTopPx);

        expect(layoutSections(opened.document.sections).length).toBe(2);
    });

    it('finds the main part through the RELATIONSHIPS, not by its name', async () => {
        const opened = await openWordFile(file('word-authored.docx'), FONTS);

        expect(opened.documentPart).toBe('word/document.xml');
        expect(opened.stylesPart).toBe('word/styles.xml');
        // 100 body paragraphs plus the document's 4 tables.
        expect(opened.document.paragraphs.length).toBe(104);
    });

    it('follows a package that names its main part something else', async () => {
        // The conventional path is a convention, not a rule. This rewrites the
        // package to use a different name and expects it to open regardless —
        // which is the only way to prove the relationship is really followed.
        const original = OpcPackage.open(file('lease-landscape.docx'));
        const relationships = await original.text('_rels/.rels');

        expect(relationships).toContain('word/document.xml');

        // Point the relationship at the styles part instead and confirm the
        // reader goes where it is told rather than where it expects.
        original.replace('_rels/.rels', relationships.replace('word/document.xml', 'word/styles.xml'));

        await expect(openWordFile(await original.save(), FONTS)).rejects.toThrow(/no <w:body>/);
    });

    it('refuses something that is not an OOXML package', async () => {
        const opc = OpcPackage.open(file('lease-landscape.docx'));
        opc.replace('_rels/.rels', '<Relationships xmlns="x"/>');

        await expect(openWordFile(await opc.save(), FONTS)).rejects.toThrow(/no officeDocument relationship/);
    });

    it('opens a document with no styles part at all', async () => {
        // Legal, and it simply has no styles. Throwing would reject a document
        // that Word opens.
        const opc = OpcPackage.open(file('lease-landscape.docx'));
        const rels = await opc.text('word/_rels/document.xml.rels');
        opc.replace(
            'word/_rels/document.xml.rels',
            rels.replace(/<Relationship[^>]*styles[^>]*\/>/, ''),
        );

        const opened = await openWordFile(await opc.save(), FONTS);

        expect(opened.stylesPart).toBeNull();
        expect(opened.document.paragraphs.length).toBeGreaterThan(0);
    });

    describe('the full edit loop', () => {
        it('opens, edits one paragraph, saves, and re-opens with the change', async () => {
            // This is what the editor will do. It has to work without damaging
            // anything else in the package.
            const original = file('lease-landscape.docx');
            const opened = await openWordFile(original, FONTS);

            const xml = XmlDocument.parse(await opened.package.text(opened.documentPart));
            xml.root.descendants('w:t')[0]!.setText('ИЗМЕНЁННЫЙ ЗАГОЛОВОК');
            opened.package.replace(opened.documentPart, xml.toString());

            const saved = await opened.package.save();
            const reopened = await openWordFile(saved, FONTS);

            expect(paragraphOf(reopened, 0).runs.map((r) => r.text).join(''))
                .toBe('ИЗМЕНЁННЫЙ ЗАГОЛОВОК');

            // Still three pages, still landscape: an edit to one run must not
            // disturb the section properties or the page breaks.
            expect(layoutPages(reopened.document.paragraphs, reopened.document.geometry).length).toBe(3);
            expect(reopened.document.geometry.widthPx).toBe(twipsToPx(16838));
        });

        it('leaves every other part of the package untouched', async () => {
            const original = file('lease-landscape.docx');
            const opened = await openWordFile(original, FONTS);

            const xml = XmlDocument.parse(await opened.package.text(opened.documentPart));
            xml.root.descendants('w:t')[0]!.setText('x');
            opened.package.replace(opened.documentPart, xml.toString());

            const before = OpcPackage.open(original);
            const after = OpcPackage.open(await opened.package.save());

            for (const name of before.names) {
                if (name === opened.documentPart) {
                    continue;
                }
                expect(await after.text(name)).toBe(await before.text(name));
            }
        });
    });
});

describe('page furniture, end to end', () => {
    // Both fixtures were printed through LibreOffice and read back page by
    // page; the expectations below ARE that output. A rule this engine got
    // wrong for a whole slice — which header an even page of a later section
    // draws — is one an assertion written from the code would have agreed with.

    it('draws the header LibreOffice draws, on every page', async () => {
        // Five pages, three sections' worth of rules interacting: `w:titlePg`
        // on both sections, `w:evenAndOddHeaders` document-wide, and section
        // two opening on document page four — its own first page AND an even
        // one. LibreOffice prints FIRST-B there.
        const opened = await openWordFile(file('title-page-parity.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages.map((page) => furnitureText(page.header)))
            .toEqual(['FIRST-A', 'EVEN-A', 'DEF-A', 'FIRST-B', 'DEF-B']);
        expect(pages.map((page) => page.pageNumber)).toEqual([1, 2, 3, 4, 5]);
    });

    it('gives a later section its EVEN header on an even page', async () => {
        // The purest form of the rule, with no `w:titlePg` to mask it: section
        // two opens on document page four, which is that section's FIRST page
        // and the document's FOURTH. LibreOffice draws EVEN-B.
        //
        // Counting parity from within the section makes page four odd, and the
        // engine drew DEFAULT-B there until this fixture was printed.
        const opened = await openWordFile(file('even-odd-sections.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages.map((page) => furnitureText(page.header)))
            .toEqual(['DEFAULT-A', 'EVEN-A', 'DEFAULT-A', 'EVEN-B']);
    });

    it('leaves the inserted parity page completely empty', async () => {
        // Section two asks to start on an odd page, so page two is inserted.
        // LibreOffice prints it with no header and no footer at all, and the
        // pages either side of it keep theirs.
        const opened = await openWordFile(file('odd-start-blank-page.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages.length).toBe(3);
        expect(pages.map((page) => furnitureText(page.header)))
            .toEqual(['HEAD-A', null, 'HEAD-B']);
        expect(pages.map((page) => furnitureText(page.footer)))
            .toEqual(['FOOT-A', null, 'FOOT-B']);
    });

    it('puts the footer at the bottom of the paper and the header at the top', async () => {
        const opened = await openWordFile(file('odd-start-blank-page.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const { heightPx } = page!.geometry;

        // Half an inch in from each edge — `w:header` and `w:footer` are both
        // 708 twips in this document.
        expect(page!.header!.topPx).toBeCloseTo(twipsToPx(708), 6);
        expect(page!.footer!.topPx + page!.footer!.heightPx)
            .toBeCloseTo(heightPx - twipsToPx(708), 6);
        // And the body sits between them.
        expect(page!.lines[0]!.yPx).toBeGreaterThan(page!.header!.topPx);
        expect(page!.lines[0]!.yPx).toBeLessThan(page!.footer!.topPx);
    });
});

describe('alignment, end to end', () => {
    /** CSS pixels to the PDF points LibreOffice reports positions in. */
    const toPoints = (px: number): number => px * 72 / 96;

    it('starts each line where LibreOffice starts it', async () => {
        // A4, one-inch margins: the column runs 72pt to 523.28pt. The numbers
        // below were read out of LibreOffice's own PDF — the `Td` operand of
        // each line's text block — not computed from this engine.
        //
        // LibreOffice insets every line by a further 0.1pt, which is why there
        // is a tolerance at all — `toBeCloseTo(x, 0)` allows half a point. The
        // measured gaps are 0.10, 0.12 and 0.01; a misread alignment would be
        // out by a hundred.
        const opened = await openWordFile(file('alignment.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const at = (index: number): number => toPoints(pages[0]!.lines[index]!.xPx);

        expect(at(0)).toBeCloseTo(72.10, 0);    // left
        expect(at(1)).toBeCloseTo(279.15, 0);   // centre
        expect(at(2)).toBeCloseTo(494.60, 0);   // right
        expect(at(3)).toBeCloseTo(72.10, 0);    // justified — the start does not move
    });

    it('stretches a justified line but not the one that ends the paragraph', async () => {
        const opened = await openWordFile(file('alignment.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const spacing = page!.lines.map((placed) => placed.wordSpacingPx);

        // Lines 3 and 4 are the two halves of the justified paragraph.
        expect(spacing[3]).toBeGreaterThan(0);
        expect(spacing[4]).toBe(0);
        // Lines 5 and 6 are a justified paragraph the AUTHOR split with a break,
        // so neither of them is stretched either.
        expect(spacing[5]).toBe(0);
        expect(spacing[6]).toBe(0);
    });

    it('fills the column exactly on the line it does stretch', async () => {
        const opened = await openWordFile(file('alignment.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const justified = page!.lines[3]!;
        const { widthPx, marginLeftPx, marginRightPx } = page!.geometry;
        const column = widthPx - marginLeftPx - marginRightPx;

        // What justification MEANS: the stretched gaps take the line to the
        // right margin. A spacing that merely looked plausible would not.
        expect(justified.line.widthPx + justified.wordSpacingPx * justified.line.spaceGaps)
            .toBeCloseTo(column, 6);
    });
});

describe('page fields, end to end', () => {
    /** The text of a placed header or footer. */
    const textOf = (placed: { lines: readonly { line: { pieces: readonly { text: string }[] } }[] } | undefined): string =>
        (placed?.lines ?? []).map((l) => l.line.pieces.map((p) => p.text).join('')).join('');

    it('answers PAGE and NUMPAGES the way LibreOffice does', async () => {
        // The fixture caches 99 for both PAGE fields and 77 for NUMPAGES, so a
        // reader that simply kept Word's cached value would print those. Printed
        // through LibreOffice the pages read "H1 / Page 1 of 3" and so on, and
        // these are those.
        //
        // The header uses the SIMPLE field form and the footer the COMPLEX one,
        // so one file covers both.
        const opened = await openWordFile(file('page-fields.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages.length).toBe(3);
        expect(pages.map((page) => textOf(page.header))).toEqual(['H1', 'H2', 'H3']);
        expect(pages.map((page) => textOf(page.footer)))
            .toEqual(['Page 1 of 3', 'Page 2 of 3', 'Page 3 of 3']);
    });

    it('keeps no trace of the cached values', async () => {
        // Stated separately because the assertion above would still pass if the
        // cache leaked in somewhere the joins hide.
        const opened = await openWordFile(file('page-fields.docx'), FONTS);
        const drawn = layoutSections(opened.document.sections)
            .flatMap((page) => [textOf(page.header), textOf(page.footer)])
            .join(' ');

        expect(drawn).not.toContain('99');
        expect(drawn).not.toContain('77');
    });
});

describe('baselines, end to end', () => {
    /** Page-top to the line's baseline, in the PDF points LibreOffice reports. */
    const baselinesOf = (pages: ReturnType<typeof layoutSections>): number[] =>
        pages[0]!.lines.map((placed) => (placed.yPx + placed.baselinePx) * 72 / 96);

    it('puts the baseline where LibreOffice puts it, under every line rule', async () => {
        // Nine paragraphs: exact at 12/24/36pt, atLeast at 12/24/36, and auto at
        // single/1.5/double, each separated by a naturally spaced line.
        //
        // The expected numbers are LibreOffice's own, read off the `Td` operand
        // of each line and converted from its bottom-left origin. They are NOT
        // computed from this engine — the three rules were fitted to them.
        //
        // LibreOffice insets every line by a constant ~0.02pt here, hence the
        // tolerance; the rules themselves differ by whole points.
        const opened = await openWordFile(file('baseline-rules.docx'), FONTS);
        const measured = baselinesOf(layoutSections(opened.document.sections));
        const height = 841.9;
        const expected = [
            760.489, 748.789, 736.989, 715.689, 701.489, 670.589, 653.989,
            641.989, 630.489, 606.489, 594.989, 558.989, 547.489,
            535.989, 524.489, 512.989, 495.739, 484.239, 461.239,
        ].map((fromBottom) => height - fromBottom);

        expect(measured.length).toBe(expected.length);
        measured.forEach((value, index) => {
            expect(value).toBeCloseTo(expected[index]!, 1);
        });
    });

    it('puts the EXACT baseline in the same place whatever the font', async () => {
        // Liberation Serif and Liberation Sans have different natural baselines,
        // and at the same exact height LibreOffice gives them the same one —
        // which is why the exact rule is a proportion rather than a metric.
        const opened = await openWordFile(file('baseline-fonts.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const lines = page!.lines;

        // Lines 0, 2 and 4 are the exact ones; 1, 3 and 5 are natural.
        expect(lines[0]!.baselinePx).toBeCloseTo(lines[2]!.baselinePx, 6);
        expect(lines[2]!.baselinePx).toBeCloseTo(lines[4]!.baselinePx, 6);
        // The natural ones do differ, or the test above would prove nothing.
        expect(lines[1]!.baselinePx).not.toBeCloseTo(lines[5]!.baselinePx, 2);
    });
});

describe('drawing a page, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    /** Only the text, since borders and shading share the operation union. */
    const textOps = (ops: readonly DrawOp[]): TextOp[] =>
        ops.filter((op): op is TextOp => 'text' === op.kind);

    it('puts the ink where LibreOffice puts it', async () => {
        // The whole chain, ending in coordinates: zip, XML, styles, fonts,
        // measurement, breaking, pagination, alignment, baselines, draw ops.
        //
        // Both numbers per line are LibreOffice's own — the `Td` operand of
        // each text block, converted from its bottom-left origin. LibreOffice
        // insets every line by ~0.1pt across and ~0.03pt down, which is the
        // whole of the tolerance; a misplaced line would be out by tens.
        const opened = await openWordFile(file('alignment.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = textOps(renderPage(page!).ops);

        const at = (index: number): [number, number] =>
            [toPoints(ops[index]!.xPx), toPoints(ops[index]!.yPx)];

        // Left, centred and right — three alignments, three baselines.
        expect(at(0)[0]).toBeCloseTo(72.10, 0);
        expect(at(0)[1]).toBeCloseTo(841.9 - 760.489, 0);
        expect(at(1)[0]).toBeCloseTo(279.15, 0);
        expect(at(1)[1]).toBeCloseTo(841.9 - 748.989, 0);
        expect(at(2)[0]).toBeCloseTo(494.60, 0);
        expect(at(2)[1]).toBeCloseTo(841.9 - 737.489, 0);
    });

    it('runs a justified line right up to the margin', async () => {
        // What justification MEANS, asserted on the drawn result rather than on
        // the number that produced it: the last word of a stretched line ends
        // at the right margin, having been placed word by word.
        const opened = await openWordFile(file('alignment.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const justified = page!.lines[3]!;
        const ops = textOps(renderPage(page!).ops)
            .filter((op) => op.yPx === justified.yPx + justified.baselinePx);

        expect(justified.wordSpacingPx).toBeGreaterThan(0);
        expect(ops.length).toBeGreaterThan(1);

        const last = ops[ops.length - 1]!;
        const right = last.xPx + last.font.measureAdvance(last.text, last.sizePx).widthPt;
        const { widthPx, marginRightPx } = page!.geometry;

        expect(right).toBeCloseTo(widthPx - marginRightPx, 6);
    });

    it('draws a header, a footer and their resolved page numbers', async () => {
        const opened = await openWordFile(file('page-fields.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const drawn = (index: number): string =>
            textOps(renderPage(pages[index]!).ops).map((op) => op.text).join('');

        // Furniture first, body last — and the fields answered per page.
        expect(drawn(0)).toBe('H1Page 1 of 3One');
        expect(drawn(2)).toBe('H3Page 3 of 3Three');
    });

    it('serialises a real document to SVG with none of its text lost', async () => {
        // Round-tripped through this package's own XML parser and compared with
        // the page — a count would pass while quietly dropping a run, and this
        // fixture is Cyrillic, so it also states that the escaping survives
        // non-Latin text.
        const opened = await openWordFile(file('lease-landscape.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const document = XmlDocument.parse(renderPageToSvg(page!));

        const drawn = document.root.elements('text').map((element) => element.text).join('');
        const expected = page!.lines
            .map((placed) => placed.line.pieces.map((piece) => piece.text).join(''))
            .join('');

        expect(document.root.name).toBe('svg');
        expect(drawn).toBe(expected);
        expect(drawn).toContain('Договор');
    });
});

describe('table borders, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('draws the rules LibreOffice draws, in weight, colour and style', async () => {
        // Each of the six table sides is declared at a DIFFERENT width and
        // colour, and one cell overrides its own left border, so no two rules
        // can be confused for one another.
        //
        // The pairs below are LibreOffice's own stroke operators for this very
        // file — `RG` for the colour, `w` for the width, `d` for the dash — read
        // out of its PDF, not computed from this engine.
        const opened = await openWordFile(file('table-borders.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const rules = renderPage(page!).ops.filter((op) => 'line' === op.kind);

        const drawn = new Set(
            rules.map((op) => `${op.color} ${toPoints(op.widthPx).toFixed(1)} ${op.style}`),
        );

        expect(drawn).toEqual(new Set([
            '#FF0000 1.0 solid',    // table top
            '#808080 0.5 dashed',   // insideH
            '#00FF00 2.0 solid',    // table left
            '#00FFFF 1.5 dotted',   // insideV
            '#FF00FF 4.0 solid',    // table right
            '#0000FF 3.0 solid',    // table bottom
            '#000000 5.0 solid',    // the cell's own override
        ]));
    });

    it('shades the cell that asked for it, and only that cell', async () => {
        const opened = await openWordFile(file('table-borders.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const fills = renderPage(page!).ops.filter((op) => 'rect' === op.kind);

        expect(fills.length).toBe(1);
        expect(fills[0]!.fill).toBe('#D9D9D9');
    });

    it('gives a shared edge to BOTH cells that meet on it', async () => {
        // Two cells side by side each draw the boundary between them, and where
        // they agree the second lands exactly on the first. LibreOffice dedupes
        // and draws it once; the difference does not show, and the conflict rule
        // for edges that DISAGREE is not implemented.
        const opened = await openWordFile(file('table-borders.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const inside = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind && '#00FFFF' === op.color);

        expect(inside.length).toBe(4);
        expect(new Set(inside.map((op) => op.x1Px)).size).toBe(1);
    });
});

describe('text colour, end to end', () => {
    it('colours each run the way LibreOffice colours it', async () => {
        // The fills below are LibreOffice's own `rg` operators for this file:
        // red, blue, black, green from a character style, and black for the run
        // that asks for `auto` OVER that style.
        const opened = await openWordFile(file('text-colour.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const drawn = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => [op.text, op.colorHex ?? '#000000']);

        expect(drawn).toEqual([
            ['Red', '#FF0000'],
            ['Blue', '#0000FF'],
            ['Plain', '#000000'],
            ['Styled', '#00FF00'],
            ['Auto', '#000000'],
        ]);
    });
});

describe('pictures, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('places the picture where LibreOffice places it', async () => {
        // The fixture's PNG is FOUR PIXELS wide and `wp:extent` asks for two
        // inches, so a reader measuring the file rather than the request lays
        // the page out completely differently.
        //
        // LibreOffice's own numbers, from the image's `cm` matrix and the `Td`
        // of the text either side of it: the picture is 144x72pt at x=104.2 with
        // its BOTTOM at the text baseline, the text resumes at 248.3, and the
        // next paragraph's baseline is 155.5pt down the page — which is only
        // possible if the line grew to hold the picture.
        const opened = await openWordFile(file('inline-image.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops;
        const picture = ops.find((op): op is ImageOp => 'image' === op.kind)!;
        const texts = ops.filter((op): op is TextOp => 'text' === op.kind);

        expect(toPoints(picture.widthPx)).toBeCloseTo(144, 1);
        expect(toPoints(picture.heightPx)).toBeCloseTo(72, 1);
        expect(toPoints(picture.xPx)).toBeCloseTo(104.2, 0);
        // Its bottom rests on the baseline of the text beside it.
        expect(toPoints(picture.yPx + picture.heightPx)).toBeCloseTo(toPoints(texts[0]!.yPx), 1);
        // The line grew: the following paragraph sits 155.5pt down.
        expect(toPoints(texts[texts.length - 1]!.yPx)).toBeCloseTo(155.5, 0);
    });

    it('embeds the picture in the SVG rather than linking a part', async () => {
        const opened = await openWordFile(file('inline-image.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const element = XmlDocument.parse(renderPageToSvg(page!)).root.element('image');
        const href = element?.attribute('href') ?? '';

        expect(href.startsWith('data:image/png;base64,')).toBe(true);
        // A real PNG, byte for byte: the signature survives the round trip.
        const decoded = Buffer.from(href.slice('data:image/png;base64,'.length), 'base64');
        expect([...decoded.subarray(0, 8)]).toEqual([137, 80, 78, 71, 13, 10, 26, 10]);
    });
});

describe('highlighting, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('fills the boxes LibreOffice fills', async () => {
        // LibreOffice's own `re f*` rectangles for this file, converted from its
        // bottom-left origin. The last one is inside an exactly spaced 24pt
        // line, which is what states that the box is the LINE and not the font.
        const opened = await openWordFile(file('highlight.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const height = toPoints(page!.geometry.heightPx);
        const fills = renderPage(page!).ops.filter((op): op is RectOp => 'rect' === op.kind);

        expect(fills.map((op) => op.fill)).toEqual(['#FFFF00', '#00FFFF', '#00FF00']);

        const box = (index: number): number[] => [
            toPoints(fills[index]!.xPx),
            height - toPoints(fills[index]!.yPx + fills[index]!.heightPx),
            toPoints(fills[index]!.heightPx),
        ];

        // x, bottom edge, height — each against LibreOffice's own.
        expect(box(0)[0]).toBeCloseTo(97, 0);
        expect(box(0)[1]).toBeCloseTo(758.44, 0);
        expect(box(0)[2]).toBeCloseTo(11.45, 0);

        expect(box(1)[0]).toBeCloseTo(94.8, 0);
        expect(box(1)[1]).toBeCloseTo(746.94, 0);

        expect(box(2)[0]).toBeCloseTo(72, 0);
        expect(box(2)[1]).toBeCloseTo(711.44, 0);
        // 24pt, not the 11.5 the font would take.
        expect(box(2)[2]).toBeCloseTo(23.95, 0);
    });

    it('leaves a run marked "none" unpainted', async () => {
        // Three fills for four paragraphs: the `w:highlight w:val="none"` one
        // draws nothing.
        const opened = await openWordFile(file('highlight.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        expect(renderPage(page!).ops.filter((op) => 'rect' === op.kind).length).toBe(3);
    });
});

describe('shared table edges, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    /**
     * The rules, with each shared edge counted once.
     *
     * Both cells that meet on an edge draw it, and after resolution they draw
     * the SAME one — so every disputed edge appears twice, identically. That is
     * deliberate: a table breaking across a page has one cell on each sheet,
     * and giving the edge to only one of them would leave the other bare.
     */
    const distinct = (rules: LineOp[]): string[] => {
        const out: string[] = [];
        for (const rule of rules) {
            const label = `${rule.color} ${toPoints(rule.widthPx).toFixed(0)}`;
            if (label !== out[out.length - 1]) {
                out.push(label);
            }
        }

        return out;
    };

    const rulesOf = async (name: string): Promise<LineOp[]> => {
        const opened = await openWordFile(file(name), FONTS);
        const [page] = layoutSections(opened.document.sections);

        return renderPage(page!).ops.filter((op): op is LineOp => 'line' === op.kind);
    };

    it('settles a disputed VERTICAL edge the way LibreOffice settles it', async () => {
        // Three pairs of neighbouring cells: different weights, a tie, and the
        // same tie with the colours swapped. LibreOffice draws 4pt red, then
        // blue, then red — which is heavier-wins with the RIGHT cell taking a
        // tie, and rules out any rule about colour.
        const rules = await rulesOf('border-conflict.docx');
        const middle = rules.filter((op) => op.x1Px === op.x2Px
            && Math.abs(toPoints(op.x1Px) - 149) < 1);

        expect(distinct(middle))
            .toEqual(['#FF0000 4', '#0000FF 2', '#008000 4', '#FF0000 2']);
    });

    it('settles a disputed HORIZONTAL edge the same way', async () => {
        // The rows disagree about the rule between them. LibreOffice draws 4pt
        // magenta, then blue, then red — heavier wins, and the LOWER row takes
        // a tie.
        const rules = await rulesOf('border-conflict-v.docx');
        const disputed = rules
            .filter((op) => op.y1Px === op.y2Px && '#C0C0C0' !== op.color);

        expect(distinct(disputed)).toEqual(['#FF00FF 4', '#0000FF 2', '#FF0000 2']);
        // And each is drawn ONCE. This used to assert six — each edge drawn
        // twice, once by each neighbour — which described what we did rather
        // than what LibreOffice does: it printed 63 rules for the 62 rows on
        // the first page of a split table, and 9 for the 8 on the second, one
        // per row plus a top on each part.
        expect(disputed.length).toBe(3);
    });

    it('insets the table by HALF its left border, as Word does', async () => {
        // Word centres the left border on the margin rather than putting the
        // table's edge there. Measured at two widths: LibreOffice draws this
        // table's 6pt left border at x=69 against a 72pt margin, and another
        // file's 2pt border at x=71.
        const rules = await rulesOf('border-conflict.docx');
        const left = rules.find((op) => '#00FF00' === op.color)!;

        expect(toPoints(left.widthPx)).toBeCloseTo(6, 1);
        expect(toPoints(left.x1Px)).toBeCloseTo(69, 0);
    });
});

describe('floating pictures, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('wraps the text the way LibreOffice wraps it', async () => {
        // A 2x1 inch picture anchored half an inch into the column, with square
        // wrapping. LibreOffice draws it with the matrix `144 0 0 72 108
        // 697.939` and starts SEVEN lines at x=261.1 — the picture's right edge
        // at 252 plus the 9pt `distR` — before line eight takes the column back
        // at 72.1.
        const opened = await openWordFile(file('floating-image.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        expect(page!.floats.length).toBe(1);
        expect(toPoints(page!.floats[0]!.xPx)).toBeCloseTo(108, 0);
        expect(toPoints(page!.floats[0]!.widthPx)).toBeCloseTo(144, 0);
        expect(toPoints(page!.floats[0]!.heightPx)).toBeCloseTo(72, 0);

        const xs = page!.lines.map((line) => Math.round(toPoints(line.xPx)));
        const wrapped = xs.filter((x) => x > 200).length;

        expect(wrapped).toBe(7);
        expect(xs[6]).toBe(261);
        expect(xs[7]).toBe(72);
    });

    it('draws the float once, at the place it was given', async () => {
        const opened = await openWordFile(file('floating-image.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const pictures = renderPage(page!).ops.filter((op): op is ImageOp => 'image' === op.kind);

        expect(pictures.length).toBe(1);
        expect(toPoints(pictures[0]!.xPx)).toBeCloseTo(108, 0);
        expect(toPoints(pictures[0]!.yPx)).toBeCloseTo(72, 0);
    });
});

describe('tab alignment, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('ranges the text against the stop the way LibreOffice does', async () => {
        // One stop three inches into the column — 288pt from the paper's edge —
        // with the SAME text after it in every paragraph, so the alignment is
        // the only thing that can move it. The numbers are LibreOffice's own.
        const opened = await openWordFile(file('tab-alignment.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim());

        // A label then the tabbed text, for each of the five paragraphs.
        const after = ops.filter((_, index) => 1 === index % 2).map((op) => toPoints(op.xPx));

        expect(after[0]).toBeCloseTo(288.10, 0);   // left: starts at the stop
        expect(after[1]).toBeCloseTo(274.80, 0);   // centre: straddles it
        expect(after[2]).toBeCloseTo(261.45, 0);   // right: ends on it
        expect(after[3]).toBeCloseTo(277.00, 0);   // decimal: separator on it
        // Too wide to fit before the stop: the tab collapses to nothing and the
        // text follows the label.
        expect(after[4]).toBeCloseTo(78.80, 0);
    });
});

describe('tab leaders, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('fills each tab the way LibreOffice fills it', async () => {
        // A contents page: a heading, a leadered tab, and a page number ranged
        // right at the margin. LibreOffice starts the dots at 127.00, the
        // hyphens at 125.90 and the underscores at 133.10 — each immediately
        // after its own heading — and puts every page number at 512.30.
        const opened = await openWordFile(file('tab-leader.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim());

        const fills = ops.filter((op) => op.text.length > 20);

        expect(fills.length).toBe(3);
        expect(toPoints(fills[0]!.xPx)).toBeCloseTo(127.00, 0);
        expect(toPoints(fills[1]!.xPx)).toBeCloseTo(125.90, 0);
        expect(toPoints(fills[2]!.xPx)).toBeCloseTo(133.10, 0);
        expect([fills[0]!.text[0], fills[1]!.text[0], fills[2]!.text[0]])
            .toEqual(['.', '-', '_']);
    });

    it('keeps the page number on the same line as its heading', async () => {
        // The stop sits at exactly the writing width, and the two are computed
        // by different routes — so an exact comparison put every page number on
        // a line of its own. Four headings, four lines.
        const opened = await openWordFile(file('tab-leader.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        expect(page!.lines.length).toBe(4);

        const numbers = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind && '12' === op.text);

        expect(numbers.length).toBe(4);
        expect(numbers.every((op) => Math.abs(toPoints(op.xPx) - 512.30) < 1)).toBe(true);
    });
});

describe('field numbering switches, end to end', () => {
    it('numbers each field the way LibreOffice numbers it', async () => {
        // A footer carrying the same PAGE field under six switches, over three
        // pages. LibreOffice prints i/ii/iii for `roman`, I/II/III for `ROMAN`,
        // a/b/c and A/B/C for the alphabetics, digits for `MERGEFORMAT` — which
        // is not a numbering switch — and `iii` for NUMPAGES on every page.
        const opened = await openWordFile(file('field-switches.docx'), FONTS);
        const drawn = layoutSections(opened.document.sections).map((page) =>
            (page.footer?.lines ?? [])
                .map((line) => line.line.pieces.map((piece) => piece.text).join(''))
                .join(''));

        expect(drawn).toEqual([
            'p1 m1 ri RI aa AA niii',
            'p2 m2 rii RII ab AB niii',
            'p3 m3 riii RIII ac AC niii',
        ]);
    });
});

describe('underlining, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('strikes through exactly where LibreOffice strikes', async () => {
        // The strikeout comes from OS/2, and LibreOffice uses the same table:
        // it draws this face 2.6pt above the baseline with a rule half a point
        // thick, over a run 36.1pt wide.
        const opened = await openWordFile(file('underline.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops;
        const strike = ops.filter((op): op is LineOp => 'line' === op.kind)[2]!;
        const baseline = page!.lines[2]!.yPx + page!.lines[2]!.baselinePx;

        expect(toPoints(baseline - strike.y1Px)).toBeCloseTo(2.6, 1);
        expect(toPoints(strike.widthPx)).toBeCloseTo(0.5, 1);
        expect(toPoints(strike.x2Px - strike.x1Px)).toBeCloseTo(36.1, 0);
    });

    it('runs each rule the width LibreOffice runs it', async () => {
        // 35.5pt at ten point and 71.1 at twenty — the run's own width, which
        // is the one thing a decoration must never get wrong.
        const opened = await openWordFile(file('underline.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const drawn = renderPage(page!).ops.filter((op): op is LineOp => 'line' === op.kind);

        expect(toPoints(drawn[0]!.x2Px - drawn[0]!.x1Px)).toBeCloseTo(35.5, 0);
        expect(toPoints(drawn[1]!.x2Px - drawn[1]!.x1Px)).toBeCloseTo(71.1, 0);
    });

    it('leaves a run marked "none" undecorated', async () => {
        // Five paragraphs, and the last asks for no underline: one rule for the
        // single, one for the strike, two for the double, and none for it.
        const opened = await openWordFile(file('underline.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const drawn = renderPage(page!).ops.filter((op) => 'line' === op.kind);

        expect(drawn.length).toBe(5);
    });
});

describe('scripts and capitals, end to end', () => {
    const toPoints = (px: number): number => px * 72 / 96;

    it('places a script exactly where LibreOffice places it', async () => {
        // The same two characters at the same nominal size in three
        // paragraphs, so only the alignment can move them. LibreOffice draws
        // both scripts at 5.8pt, the superscript's baseline 3.95pt above the
        // line's and the subscript's 0.9 below.
        const opened = await openWordFile(file('vert-align.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim());

        const height = 841.9;

        expect(toPoints(ops[1]!.sizePx)).toBeCloseTo(5.8, 1);
        expect(toPoints(ops[1]!.yPx)).toBeCloseTo(height - 764.439, 1);
        expect(toPoints(ops[4]!.yPx)).toBeCloseTo(height - 748.089, 1);
        // And an explicit `baseline` leaves the run alone.
        expect(toPoints(ops[7]!.sizePx)).toBeCloseTo(10, 1);
        expect(toPoints(ops[7]!.yPx)).toBeCloseTo(height - 737.489, 1);
    });

    it('capitalises the run LibreOffice capitalises', async () => {
        const opened = await openWordFile(file('vert-align.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim());

        expect(ops[10]!.text).toBe('SHOUT');
        // At the same place LibreOffice puts it, so the CAPITALS were measured.
        expect(toPoints(ops[10]!.xPx)).toBeCloseTo(96, 0);
    });
});

describe('keeping a paragraph with the next one', () => {
    /**
     * Everything a page holds, in reading order.
     *
     * Table rows and text lines have to be merged by their height on the page:
     * pagination interleaves them, and `page.lines` alone would show a heading
     * without the table it was kept with.
     */
    const contentsOf = (page: Page): string[] => [
        ...page.lines.map((placed) => ({
            yPx: placed.yPx,
            text: placed.line.pieces.map((piece) => piece.text).join(''),
        })),
        ...page.rows.map((row) => ({
            yPx: row.yPx,
            text: row.cells
                .flatMap((cell) => cell.lines.map(
                    (line) => line.line.pieces.map((piece) => piece.text).join('')))
                .join(' '),
        })),
    ]
        .sort((one, other) => one.yPx - other.yPx)
        .map((entry) => entry.text)
        .filter((text) => '' !== text);

    it('paginates keep-next.docx the way LibreOffice does', async () => {
        // Fourteen lines to a page, and each group begins on a page of its own
        // so they cannot interfere. Printed through LibreOffice, this is
        // exactly where the text falls.
        const opened = await openWordFile(file('keep-next.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages.map(contentsOf)).toEqual([
            // The chain moved BOTH headings down, not just the lower one.
            ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L10', 'L11', 'L12'],
            ['H1', 'H2', 'BODY'],
            // A heading is kept with the table below it.
            ['M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12',
                'M13'],
            ['HEAD', 'R1A R1B', 'R2A R2B'],
            // Nothing follows the last one, so it stays where it fell.
            ['N1', 'N2', 'N3', 'N4', 'N5', 'N6', 'N7', 'N8', 'N9', 'N10', 'N11', 'N12',
                'N13', 'LAST'],
        ]);
    });

    it('reads w:keepNext, and leaves it off a paragraph without it', async () => {
        const opened = await openWordFile(file('keep-next.docx'), FONTS);

        expect(paragraphOf(opened, 12).style.keepWithNext).toBe(true);
        expect(paragraphOf(opened, 13).style.keepWithNext).toBe(true);
        expect(paragraphOf(opened, 14).style.keepWithNext).toBe(false);
        expect(paragraphOf(opened, 28).style.keepWithNext).toBe(true);
    });
});

describe('vertical alignment and row height, end to end', () => {
    // A4 is 841.9pt tall and the PDF measures up from its foot, so this is
    // what each of LibreOffice's baselines becomes here.
    const AT = (fromFoot: number): number => 841.9 - fromFoot;

    const baselines = async (): Promise<Map<string, number>> => {
        const opened = await openWordFile(file('cell-vertical.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const found = new Map<string, number>();
        for (const op of renderPage(page!).ops) {
            if ('text' === op.kind && '' !== op.text.trim()) {
                found.set(op.text.trim(), op.yPx * 72 / 96);
            }
        }

        return found;
    };

    it('puts every cell where LibreOffice puts it', async () => {
        const y = await baselines();
        const at = (text: string): number => {
            const found = y.get(text);
            if (undefined === found) {
                throw new Error(`nothing drawn for ${text}`);
            }

            return found;
        };

        // Top by default, and top when it says so.
        expect(at('Atop')).toBeCloseTo(AT(775.79), 1);
        expect(at('Atop')).toBeCloseTo(at('A1'), 6);

        // Centred lands on the middle line of the three beside it, bottomed on
        // the last — relative to the neighbour, because a shift that moved
        // BOTH cells would keep the absolute figure right and the layout wrong.
        expect(at('Bctr')).toBeCloseTo(AT(729.79), 1);
        expect(at('Bctr')).toBeCloseTo(at('B2'), 6);
        expect(at('Cbot')).toBeCloseTo(AT(683.79), 1);
        expect(at('Cbot')).toBeCloseTo(at('C3'), 6);

        // Bottomed in a row made tall by trHeight rather than by a neighbour:
        // one inch of row, one line of text, so 60.5pt of it above the text.
        expect(at('Dbot')).toBeCloseTo(AT(611.79), 1);
        expect(at('Dbot') - at('d')).toBeCloseTo(60.5, 1);

        // A height with no rule at all is still a minimum: the row is the inch
        // it asked for and the next row starts below it.
        expect(at('Eauto') - at('d')).toBeCloseTo(72, 1);
        // A minimum SMALLER than the text grows to the text: three lines.
        expect(at('F1') - at('Eauto')).toBeCloseTo(72, 1);
        expect(at('G1') - at('F1')).toBeCloseTo(34.5, 1);
        // ...where `exact` does not, and keeps its 15pt whatever the text.
        expect(at('LAST') - at('G1')).toBeCloseTo(15, 1);
    });

    it('drops the line that hangs out of an exact row', async () => {
        // Three lines in a row 15pt tall. LibreOffice drew two and dropped the
        // third, whose top had fallen past the foot of the row.
        const y = await baselines();

        expect(y.has('G1')).toBe(true);
        expect(y.has('G2')).toBe(true);
        expect(y.has('G3')).toBe(false);
    });

    it('reads w:vAlign and w:trHeight off the file', async () => {
        const opened = await openWordFile(file('cell-vertical.docx'), FONTS);
        const table = opened.document.paragraphs.find(isTable)!;

        expect(table.rows[0]!.cells[1]!.verticalAlign).toBe('top');
        expect(table.rows[3]!.cells[1]!.verticalAlign).toBeUndefined();
        expect(table.rows[1]!.cells[1]!.verticalAlign).toBe('center');
        expect(table.rows[2]!.cells[1]!.verticalAlign).toBe('bottom');
        expect(table.rows[4]!.heightRule).toBe('atLeast');
        expect(table.rows[6]!.heightRule).toBe('exact');
        expect(table.rows[7]!.heightPx).toBeUndefined();
    });
});

describe('vertical merging, end to end', () => {
    const AT = (fromFoot: number): number => 841.9 - fromFoot;

    const drawn = async (): Promise<Map<string, number>> => {
        const opened = await openWordFile(file('vertical-merge.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const found = new Map<string, number>();
        for (const op of renderPage(page!).ops) {
            if ('text' === op.kind && '' !== op.text.trim()) {
                found.set(op.text.trim(), op.yPx * 72 / 96);
            }
        }

        return found;
    };

    it('places every merged cell where LibreOffice places it', async () => {
        const y = await drawn();
        const at = (text: string): number => {
            const found = y.get(text);
            if (undefined === found) {
                throw new Error(`nothing drawn for ${text}`);
            }

            return found;
        };

        // The top of a three-row span.
        expect(at('TOP')).toBeCloseTo(AT(775.79), 1);
        expect(at('TOP')).toBeCloseTo(at('a1'), 6);

        // Centred over the span lands on the middle row of the three, which is
        // a row the cell does not itself belong to.
        expect(at('CTR')).toBeCloseTo(AT(729.79), 1);
        expect(at('CTR')).toBeCloseTo(at('b2'), 6);

        // Five merged lines across two one-line rows: the first row keeps its
        // line and the second grows to four, so the row after starts 57.5pt
        // below the top of the span rather than 23.
        expect(at('T1')).toBeCloseTo(AT(706.79), 1);
        expect(at('LAST') - at('T1')).toBeCloseTo(57.5, 1);
        expect(at('c2') - at('c1')).toBeCloseTo(11.5, 1);
    });

    it('draws nothing at all for a swallowed cell', async () => {
        const y = await drawn();

        expect(y.has('DROPPED')).toBe(false);
    });

    it('gives the merged cell one box over its whole span', async () => {
        // What keeps a border from being drawn through the middle of it.
        const opened = await openWordFile(file('vertical-merge.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const rows = page!.rows;

        expect(rows[0]!.cells[0]!.heightPx).toBeCloseTo(rows[0]!.heightPx * 3, 6);
        expect(rows[1]!.cells.length).toBe(1);
        expect(rows[2]!.cells.length).toBe(1);
    });

    it('reads w:vMerge, with a bare element meaning continue', async () => {
        const opened = await openWordFile(file('vertical-merge.docx'), FONTS);
        const table = opened.document.paragraphs.find(isTable)!;

        expect(table.rows[0]!.cells[0]!.verticalMerge).toBe('restart');
        expect(table.rows[1]!.cells[0]!.verticalMerge).toBe('continue');
        expect(table.rows[0]!.cells[1]!.verticalMerge).toBeUndefined();
        expect(verticalSpans(table)[0]).toEqual([3, 1]);
    });
});

describe('nested tables, end to end', () => {
    const AT = (fromFoot: number): number => 841.9 - fromFoot;

    const drawn = async (): Promise<TextOp[]> => {
        const opened = await openWordFile(file('nested-table.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        return renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim());
    };

    it('places a nested table where LibreOffice places it', async () => {
        const ops = await drawn();
        const pt = (op: TextOp | undefined): { x: number; y: number } => ({
            x: (op?.xPx ?? 0) * 72 / 96,
            y: (op?.yPx ?? 0) * 72 / 96,
        });
        const at = (text: string, nth = 0) =>
            pt(ops.filter((op) => op.text.trim() === text)[nth]);

        // A line, the nested table's two rows, then a line — all in one cell.
        expect(at('above').y).toBeCloseTo(AT(775.79), 1);
        expect(at('i1').y).toBeCloseTo(AT(764.29), 1);
        expect(at('i3').y).toBeCloseTo(AT(752.79), 1);
        expect(at('below').y).toBeCloseTo(AT(741.29), 1);
        // The cell beside it is unmoved.
        expect(at('side').y).toBeCloseTo(AT(775.79), 1);

        // The nested table starts at the HOLDING cell's content origin, so its
        // own cell margin is the whole of the step in — the same margin the
        // outer table put between the page edge and 'above'.
        //
        // Not the absolute figure: this uses Word's documented 0.08in default
        // cell margin and LibreOffice uses 0.19cm, which is a quarter point per
        // level and so half a point by the time it has nested once. The step is
        // what the nesting decides; the margin it steps by is a separate
        // question, and the same one at every depth.
        const PAGE_MARGIN_PT = 56.7;

        expect(at('i1').x - at('above').x).toBeCloseTo(at('above').x - PAGE_MARGIN_PT, 1);
        // The inner column's own width, which no default is involved in.
        expect(at('i2').x - at('i1').x).toBeCloseTo(100, 1);

        // The second row's cell is NOTHING but a nested table.
        expect(at('i1', 1).y).toBeCloseTo(AT(729.79), 1);
        expect(at('s2').y).toBeCloseTo(AT(729.79), 1);
    });

    it('gives the mandatory trailing paragraph no height of its own', async () => {
        // LibreOffice draws it a twentieth of a point tall and this draws it at
        // nothing, which is the whole of the difference: the row below starts
        // 23pt after the nested table rather than 34.5.
        const ops = await drawn();
        const after = ops.find((op) => 'AFTER' === op.text.trim());

        expect((after?.yPx ?? 0) * 72 / 96).toBeCloseTo(AT(706.74), 0);
        expect((after?.yPx ?? 0) * 72 / 96 - AT(729.79)).toBeCloseTo(23, 1);
    });

    it('reads the nested table rather than reporting it away', async () => {
        const opened = await openWordFile(file('nested-table.docx'), FONTS);
        const outer = opened.document.paragraphs.find(isTable)!;
        const held = outer.rows[0]!.cells[0]!.paragraphs;

        expect(opened.document.diagnostics.map((entry) => entry.detail))
            .not.toContain('a table nested inside a cell was skipped');
        // Paragraph, table, paragraph — in the order the file has them.
        expect(held.map(isTable)).toEqual([false, true, false]);
    });
});

describe('columns, end to end', () => {
    it('flows columns.docx the way LibreOffice does', async () => {
        // Fourteen lines to a column and two columns, so twenty-eight before
        // the page turns. Printed through LibreOffice, this is where it falls.
        const opened = await openWordFile(file('columns.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const textOf = (page: Page): string[] => page.lines.map(
            (placed) => placed.line.pieces.map((piece) => piece.text).join('').trim());

        expect(pages.length).toBe(2);
        expect(textOf(pages[0]!).length).toBe(28);
        expect(textOf(pages[0]!)[0]).toBe('L1');
        expect(textOf(pages[0]!)[14]).toBe('L15');
        expect(textOf(pages[1]!)).toEqual(['L29', 'L30', 'L31', 'L32', 'L33', 'L34']);
    });

    it('puts the second column where LibreOffice puts it', async () => {
        const opened = await openWordFile(file('columns.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const page = pages[0];
        const at = (index: number): number => (page!.lines[index]?.xPx ?? 0) * 72 / 96;

        // LibreOffice printed 28.45 and 315.45 — it insets every line by a
        // tenth of a point, which is why the STEP across is asserted rather
        // than the two edges: a 251.6pt column and a 35.4pt gap, and the inset
        // cancels itself out of the difference.
        expect(at(14) - at(0)).toBeCloseTo(287, 1);
        expect(at(0)).toBeCloseTo(28.35, 0);

        // ...and the second page opens back in the LEFT column, where
        // LibreOffice printed L29.
        const second = (pages[1]?.lines[0]?.xPx ?? 0) * 72 / 96;

        expect(second).toBeCloseTo(at(0), 6);
    });
});

describe('footnotes, end to end', () => {
    it('paginates footnotes.docx the way LibreOffice does', async () => {
        // Eleven body lines fit beside one note; the twelfth carries a SECOND
        // reference, which no longer fits and takes its note to the next page.
        const opened = await openWordFile(file('footnotes.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const textOf = (page: Page): string[] => page.lines.map(
            (placed) => placed.line.pieces.map((piece) => piece.text).join('').trim());

        expect(pages.length).toBe(2);
        expect(textOf(pages[0]!)).toEqual([
            'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'mid1', 'after',
        ]);
        expect(textOf(pages[1]!)).toEqual(['end2']);
    });

    it('puts the notes and the rule where LibreOffice puts them', async () => {
        const opened = await openWordFile(file('footnotes.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const pt = (px: number): number => px * 72 / 96;
        // The PDF measures up from the foot of a 226.8pt page.
        const AT = (fromFoot: number): number => 226.8 - fromFoot;

        const first = pages[0]!.footnotes!;

        expect(pt(first.separatorYPx)).toBeCloseTo(AT(43.949), 1);
        expect(pt(first.separatorWidthPx)).toBeCloseTo(144, 6);
        expect(pt(first.separatorLeftPx)).toBeCloseTo(28.35, 1);

        // The note's own baseline, where LibreOffice drew ' first note'.
        const baseline = first.lines[0]!.yPx + first.lines[0]!.baselinePx;

        expect(pt(baseline)).toBeCloseTo(AT(30.45), 1);
        expect(first.lines[0]!.line.pieces.map((piece) => piece.text).join(''))
            .toBe('1 first note');
    });

    it('numbers the notes by the order the BODY refers to them', async () => {
        const opened = await openWordFile(file('footnotes.docx'), FONTS);
        const marks = opened.document.paragraphs
            .filter((block): block is Paragraph => !isTable(block))
            .flatMap((block) => block.runs)
            .filter((run) => undefined !== run.footnoteId);

        expect(marks.map((run) => [run.footnoteId, run.text])).toEqual([[1, '1'], [2, '2']]);
        expect([...opened.document.footnotes.keys()]).toEqual([1, 2]);
    });

    it('draws the rule and the note text', async () => {
        const opened = await openWordFile(file('footnotes.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops;
        const rules = ops.filter((op): op is LineOp => 'line' === op.kind);

        expect(rules.some((op) => op.y1Px === op.y2Px
            && op.x2Px - op.x1Px === page!.footnotes!.separatorWidthPx)).toBe(true);

        // And it is a HAIRLINE. The print draws this rule from
        // 28.350 to 172.350 — two inches exactly, which the test above pins —
        // at a width of 0.100, the thinnest LibreOffice will draw. This was a
        // whole pixel, 0.75pt: seven and a half times too heavy for a rule
        // whose only job is to separate. Its place, its length and its height
        // above the notes had all been measured; its thickness never had.
        const separator = rules.find((op) => op.y1Px === op.y2Px
            && op.x2Px - op.x1Px === page!.footnotes!.separatorWidthPx)!;

        expect(Math.round(separator.widthPx * 72 / 96 * 1000) / 1000).toBe(0.1);
        // Joined, because a line reaches the renderer as its PIECES: the note
        // arrives as its mark, its space and its two words.
        expect(ops.filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => op.text).join('')).toContain('first note');
    });
});

describe('VML pictures, end to end', () => {
    it('lays vml-picture.docx out the way LibreOffice does', async () => {
        // Four paragraphs, three of them a two-inch picture stated in a
        // different unit each time, and one with a half-inch picture between
        // two letters. LibreOffice drew A and B 42.7pt apart and put AFTER a
        // line below them.
        const opened = await openWordFile(file('vml-picture.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        expect(opened.document.diagnostics.map((entry) => entry.detail))
            .not.toContain('a VML picture (w:pict) was skipped');

        // The last picture sits BETWEEN the two letters, taking the width it
        // asked for rather than the width of the character standing in for it.
        const inline = page!.lines[page!.lines.length - 2]!.line.pieces;

        // The middle piece is the object-replacement character the picture
        // stands in as, so a caret can still be put either side of it.
        expect(inline.map((piece) => piece.text)).toEqual(['A', '￼', 'B']);
        expect(inline[1]!.widthPx).toBe(48);

        // Three pictures of the same size, whatever unit each was written in.
        const heights = page!.lines
            .flatMap((placed) => placed.line.pieces)
            .filter((piece) => undefined !== piece.image)
            .map((piece) => piece.image!.heightPx);

        expect(heights.slice(0, 3)).toEqual([96, 96, 96]);
        expect(heights[3]).toBe(48);
    });
});

describe('small caps, end to end', () => {
    it('sets small-caps.docx to the widths LibreOffice measured', async () => {
        const opened = await openWordFile(file('small-caps.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const pt = (px: number): number => px * 72 / 96;

        // Each line ends with a marker, so the width of what precedes it is
        // the width of the line's other pieces.
        const widthOf = (index: number): number => pt(page!.lines[index]!.line.pieces
            .filter((piece) => !piece.text.startsWith('|'))
            .reduce((sum, piece) => sum + piece.widthPx, 0));

        // Plain lower case, then the same letters already capital.
        expect(widthOf(0)).toBeCloseTo(30.02, 1);
        expect(widthOf(1)).toBeCloseTo(40.55, 1);
        // Lower case in small caps: the capitals at four fifths of the size.
        expect(widthOf(2)).toBeCloseTo(32.40, 1);
        // A mixture, which LibreOffice drew as a 6.65pt `A` and 27.10 after it.
        expect(widthOf(3)).toBeCloseTo(33.75, 1);
        // ...and capitals already, which small caps leaves at full width.
        expect(widthOf(4)).toBeCloseTo(40.55, 1);
    });

    it('splits the mixed run where LibreOffice split it', async () => {
        const opened = await openWordFile(file('small-caps.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const pieces = page!.lines[3]!.line.pieces;

        expect(pieces.map((piece) => piece.text)).toEqual(['A', 'BCDEF', '|d']);
        expect(pieces[0]!.widthPx * 72 / 96).toBeCloseTo(6.65, 1);
        expect(pieces[1]!.widthPx * 72 / 96).toBeCloseTo(27.10, 1);
    });
});

describe('character spacing, end to end', () => {
    it('sets letter-spacing.docx to the widths LibreOffice measured', async () => {
        // Ten of one glyph, so the base width is ten advances and the rest is
        // tracking. LibreOffice printed 50pt bare, then 55, 60, 70 and 100.
        const opened = await openWordFile(file('letter-spacing.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const widthOf = (index: number): number => page!.lines[index]!.line.pieces
            .filter((piece) => !piece.text.startsWith('|'))
            .reduce((sum, piece) => sum + piece.widthPx, 0) * 72 / 96;

        expect(widthOf(0)).toBeCloseTo(50, 1);
        expect(widthOf(1)).toBeCloseTo(55, 1);
        expect(widthOf(2)).toBeCloseTo(60, 1);
        expect(widthOf(3)).toBeCloseTo(70, 1);
        expect(widthOf(4)).toBeCloseTo(100, 1);
    });

    it('tracks the run that asked for it and no other', async () => {
        // Two runs on one line, only the first spaced. LibreOffice started the
        // second at 70pt along, which is the first run's tracked width.
        const opened = await openWordFile(file('letter-spacing.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const pieces = page!.lines[5]!.line.pieces;

        expect(pieces[0]!.widthPx * 72 / 96).toBeCloseTo(70, 1);
        expect(pieces[1]!.widthPx * 72 / 96).toBeCloseTo(50, 1);
    });
});

describe('endnotes, end to end', () => {
    it('lays endnotes.docx out the way LibreOffice does', async () => {
        // Two references, then the body's last line, then the notes — all in
        // the ordinary flow at the end of the document.
        const opened = await openWordFile(file('endnotes.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const textOf = (index: number): string => page!.lines[index]!.line.pieces
            .map((piece) => piece.text).join('');

        expect(textOf(0)).toBe('alphai');
        expect(textOf(1)).toBe('betaii');
        expect(textOf(2)).toBe('LAST');
        expect(textOf(3)).toBe('i note one');
        expect(textOf(4)).toBe('ii note two');
    });

    it('raises the mark in the text and leaves the note alone', async () => {
        // LibreOffice drew the reference 3.95pt above the line it sits on and
        // the note's own opening mark ON its baseline.
        const opened = await openWordFile(file('endnotes.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const ops = renderPage(page!).ops
            .filter((op): op is TextOp => 'text' === op.kind);
        const at = (text: string, nth = 0): number =>
            ops.filter((op) => op.text === text)[nth]!.yPx * 72 / 96;

        expect(at('alpha') - at('i')).toBeCloseTo(3.95, 1);
        // The note's own mark is the SECOND `i` drawn, and it sits on the
        // baseline of the words beside it rather than above them.
        expect(at('i', 1)).toBeCloseTo(at('note '), 6);
    });
});

describe('paragraph borders, end to end', () => {
    // Every number below was read out of the PDF LibreOffice printed from
    // THIS fixture, off the content stream's `re` and `m … l` operators —
    // rules are invisible to text extraction, so the geometry has to come from
    // the drawing operators themselves.
    //
    // A4 is 841.861pt tall and the margins are 567 twips (28.35pt), so the
    // content box runs from x=28.35 to x=566.93. LibreOffice insets what it
    // draws by around a tenth of a point, which is why these compare within
    // 0.15 rather than exactly.
    const PAGE_HEIGHT_PT = 841.861;
    const TOLERANCE_PT = 0.15;

    const pt = (px: number): number => px * 72 / 96;

    /** The four edges of a box, in points, measured from the TOP of the page. */
    const edgesOf = (box: PlacedParagraphBorder): number[] =>
        [pt(box.leftPx), pt(box.rightPx), pt(box.topPx), pt(box.bottomPx)];

    /** The same, as LibreOffice printed them — its y counts from the FOOT. */
    const printed = (
        left: number, right: number, top: number, bottom: number,
    ): number[] => [left, right, PAGE_HEIGHT_PT - top, PAGE_HEIGHT_PT - bottom];

    const boxes = async (): Promise<readonly PlacedParagraphBorder[]> => {
        const opened = await openWordFile(file('paragraph-borders.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        return page!.paragraphBorders;
    };

    const near = (ours: number[], theirs: number[]): void => {
        ours.forEach((value, index) => {
            expect(Math.abs(value - theirs[index]!)).toBeLessThan(TOLERANCE_PT);
        });
    };

    it('draws three boxes where LibreOffice drew three outlines', async () => {
        // Four bordered paragraphs, but the last two carry the same border and
        // came out as ONE outline with no rule between them.
        expect(await boxes()).toHaveLength(3);
    });

    it('puts the plain box where LibreOffice put it', async () => {
        near(edgesOf((await boxes())[0]!), printed(27.9, 567.45, 801.489, 789.039));
    });

    it('pushes the spaced box out by the six points it asked for', async () => {
        // 21.9 and 573.45 against the plain box's 27.9 and 567.45: six points
        // further out on each side, and the same six above and below.
        near(edgesOf((await boxes())[1]!), printed(21.9, 573.45, 776.489, 752.039));
    });

    it('runs one box round the pair that share a border', async () => {
        // Two lines inside one outline: 739.489 down to 715.539, which is two
        // 11.5pt line boxes with half a rule outside each end.
        near(edgesOf((await boxes())[2]!), printed(27.9, 567.45, 739.489, 715.539));
    });

    it('rules between two paragraphs of one box, where w:between asks for it', async () => {
        // A three-point `w:between` inside a one-point box. LibreOffice drew
        // the outline once — top at 801.489, bottom at 774.539 — and a single
        // rule across it at 787.989, which is halfway through the three points
        // it took out of the flow. The step from `one` to `two` was 14.5pt
        // against the 11.5 an unruled pair steps.
        const opened = await openWordFile(file('paragraph-border-between.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.paragraphBorders;

        expect(page!.paragraphBorders).toHaveLength(1);
        near(edgesOf(box!), printed(27.9, 567.45, 801.489, 774.539));
        expect(box!.innerYPx.map(pt)).toHaveLength(1);
        expect(Math.abs(pt(box!.innerYPx[0]!) - (PAGE_HEIGHT_PT - 787.989)))
            .toBeLessThan(TOLERANCE_PT);

        const step = pt(page!.lines[2]!.yPx - page!.lines[1]!.yPx);
        expect(Math.abs(step - 14.5)).toBeLessThan(TOLERANCE_PT);
    });

    it('leaves the text exactly where LibreOffice left it', async () => {
        // The box moves the paragraphs down — 12.5pt where an unbordered pair
        // steps 11.5, and 18.5 with six points of space — but never sideways.
        const opened = await openWordFile(file('paragraph-borders.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const steps = page!.lines.slice(1).map(
            (line, index) => pt(line.yPx - page!.lines[index]!.yPx));

        // above→boxed, boxed→below, below→padded, padded→after, after→first,
        // first→second, second→end.
        [12.5, 12.5, 18.5, 18.5, 12.5, 11.5, 12.5].forEach((printedStep, index) => {
            expect(Math.abs(steps[index]! - printedStep)).toBeLessThan(TOLERANCE_PT);
        });
    });
});

describe('table width, indent and alignment, end to end', () => {
    // Read off the PDF LibreOffice printed from these fixtures, from the
    // content stream's own line operators — the numbers are the CENTRES of the
    // one-point rules the tables draw, which is where our own edges sit too.
    // The text column runs from 28.35pt to 566.93 on A4 with 567-twip margins,
    // and every table states a 115-twip cell margin so both engines measure
    // the indent against the same inset.
    const TOLERANCE_PT = 0.12;
    const pt = (px: number): number => px * 72 / 96;

    /** A table's left and right edges, in points. */
    const edgesOf = (page: Page, index: number): [number, number] => {
        const cells = page.rows[index]!.cells;
        const last = cells[cells.length - 1]!;

        return [pt(cells[0]!.xPx), pt(last.xPx + last.widthPx)];
    };

    const near = (ours: readonly number[], theirs: readonly number[]): void => {
        ours.forEach((value, index) => {
            expect(Math.abs(value - theirs[index]!)).toBeLessThan(TOLERANCE_PT);
        });
    };

    const pageOf = async (name: string): Promise<Page> => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections)[0]!;
    };

    it('sizes a table with no w:tblW from its grid alone', async () => {
        // The control: 4000 + 4000 twips is 400pt, and LibreOffice printed
        // 27.9 to 427.9 — the left border hanging half its width outside the
        // margin, which is what a table starting AT the margin does.
        near(edgesOf(await pageOf('table-width.docx'), 0), [27.9, 427.9]);
    });

    it('resolves a percentage width against the text column', async () => {
        // 2500 fiftieths is half, and half of a 538.58pt column is 269.29 —
        // which is exactly what 297.2 minus 27.9 comes to.
        near(edgesOf(await pageOf('table-width.docx'), 1), [27.9, 297.2]);
    });

    it('centres a narrow table in the column', async () => {
        near(edgesOf(await pageOf('table-width.docx'), 2), [163.0, 432.3]);
    });

    it('pushes a narrow table to the right margin', async () => {
        near(edgesOf(await pageOf('table-width.docx'), 3), [297.7, 567.0]);
    });

    it('indents to the CELL TEXT, so the border lands a cell margin further left', async () => {
        // Half an inch of `w:tblInd` against a 5.75pt cell margin: LibreOffice
        // put the border at 58.6 — 30.25pt in, not 36 — and the text at the
        // full 64.45, which IS 28.35 plus 36. The hang is gone with it.
        near(edgesOf(await pageOf('table-width.docx'), 4), [58.6, 458.7]);
    });

    it('sizes a table stated in twips, scaling the whole grid to fit', async () => {
        // 4000 twips is 200pt, and the two 4000-twip columns came out 100
        // apiece rather than one being trimmed.
        const page = await pageOf('table-width.docx');

        near(edgesOf(page, 5), [27.9, 227.9]);
        expect(page.rows[5]!.cells.map((cell) => Math.round(pt(cell.widthPx))))
            .toEqual([100, 100]);
    });

    it('resolves a percentage COLUMN against the table, not the page', async () => {
        // Two fifty-percent columns inside a table declared 6000 twips wide
        // printed 150pt each. Against the page they would have been 269.
        const page = await pageOf('table-width.docx');

        near(edgesOf(page, 6), [27.9, 327.9]);
        expect(page.rows[6]!.cells.map((cell) => Math.round(pt(cell.widthPx))))
            .toEqual([150, 150]);
    });

    it('leaves percentage columns alone when the table declares no width', async () => {
        // Nothing to resolve them against. LibreOffice printed the GRID —
        // 27.9 to 427.9, the untouched 400pt — and so does this.
        near(edgesOf(await pageOf('table-width-edges.docx'), 0), [27.9, 427.9]);
    });

    it('drops the indent once the table is centred', async () => {
        // Half an inch of `w:tblInd` AND a centring: LibreOffice printed it in
        // exactly the place a centred table with no indent prints.
        const indented = edgesOf(await pageOf('table-width-edges.docx'), 1);

        near(indented, [163.0, 432.3]);
        near(indented, edgesOf(await pageOf('table-width.docx'), 2));
    });

    it('keeps the cell-margin pull for an indent it cannot resolve', async () => {
        // A percentage indent came to nothing in LibreOffice — but the table
        // still moved LEFT by the cell margin, to 22.6 from 27.9, putting its
        // text on the margin. Present and unresolvable is nought, not absent.
        near(edgesOf(await pageOf('table-width-edges.docx'), 2), [22.6, 422.7]);
    });
});

describe('w:tcMar, end to end', () => {
    // Off the PDF again. Every table states a 115-twip (5.75pt) cell margin,
    // so each override is read against a known control on the same page — and
    // the heights are compared as DIFFERENCES from that control, which is what
    // makes them independent of whose font metrics decide the line.
    const TOLERANCE_PT = 0.12;
    const pt = (px: number): number => px * 72 / 96;

    const pageOf = async (name: string): Promise<Page> => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections)[0]!;
    };

    const near = (ours: number, theirs: number): void => {
        expect(Math.abs(ours - theirs)).toBeLessThan(TOLERANCE_PT);
    };

    it('puts the text a stated LEFT margin in from the cell edge', async () => {
        // 500 twips is 25pt, against the table's own 5.75. LibreOffice drew
        // the control at 33.70 and the override at 52.95, off a cell edge at
        // 27.9 either way.
        const page = await pageOf('cell-margins.docx');

        near(pt(page.rows[0]!.cells[0]!.lines[0]!.xPx), 33.70);
        near(pt(page.rows[1]!.cells[0]!.lines[0]!.xPx), 52.95);
    });

    it('grows the ROW by a stated top margin and pushes the text down it', async () => {
        // 400 twips is 20pt. LibreOffice's row went from 12.5pt tall to 32.5,
        // and the text moved down the whole 20 with it.
        const page = await pageOf('cell-margins.docx');
        const control = page.rows[0]!;
        const padded = page.rows[1]!;

        near(pt(padded.heightPx - control.heightPx), 20);
        near(
            pt((padded.cells[0]!.lines[0]!.yPx - padded.yPx)
                - (control.cells[0]!.lines[0]!.yPx - control.yPx)),
            20,
        );
    });

    it('grows the row by a stated bottom margin and leaves the text alone', async () => {
        // 800 twips is 40pt: LibreOffice's row went from 12.5pt to 52.5, and
        // the text stayed at the top of it where the control put it. The
        // control is the other fixture's first row — a single line under the
        // same table margins, which is what makes the two comparable.
        const control = (await pageOf('cell-margins.docx')).rows[0]!;
        const padded = (await pageOf('cell-margins-edges.docx')).rows[2]!;

        near(pt(padded.heightPx - control.heightPx), 40);
        near(
            pt(padded.cells[0]!.lines[0]!.yPx - padded.yPx),
            pt(control.cells[0]!.lines[0]!.yPx - control.yPx),
        );
    });

    it('narrows the text column by a stated right margin', async () => {
        // The control wrapped 'wwwww wwwww wwwww wwwww' onto three lines in a
        // 100pt column; 40pt of right margin left room for one word a line.
        const page = await pageOf('cell-margins-edges.docx');

        expect(page.rows[0]!.cells[0]!.lines).toHaveLength(3);
        expect(page.rows[1]!.cells[0]!.lines).toHaveLength(4);
    });

    it('leaves the cell beside it on the table’s own margin', async () => {
        // The override is the CELL'S. LibreOffice put the second cell's text
        // at 252.95 when that cell declared 500 twips, and its neighbour stayed
        // at 33.70 — the table's 5.75 off a cell edge of 27.9.
        const page = await pageOf('cell-margins-edges.docx');
        const row = page.rows[3]!;

        near(pt(row.cells[0]!.lines[0]!.xPx), 33.70);
        near(pt(row.cells[1]!.lines[0]!.xPx), 252.95);
    });
});

describe('w:position, end to end', () => {
    // Off the PDF: each line holds a plain run, the run under test, and a
    // plain run again, so the baseline either side is on the page to measure
    // the raise against. The font is 10pt Liberation Serif, whose line stands
    // 1.15 times its size.
    const TOLERANCE_PT = 0.12;
    const pt = (px: number): number => px * 72 / 96;

    /** Each line's pieces: the plain run, the one under test, then plain again. */
    const drawn = async (name: string): Promise<LinePiece[][]> => {
        const opened = await openWordFile(file(name), FONTS);
        const [page] = layoutSections(opened.document.sections);

        return page!.lines.map((placed) => [...placed.line.pieces]);
    };

    /**
     * How far the middle run sits ABOVE the plain run beside it, in points.
     *
     * Straight off the shift, because that is what the renderer adds to the
     * line's own baseline — `display-list` draws at `baseline + shift`, so a
     * piece with no shift is the line and the difference is the whole story.
     */
    const riseOf = (pieces: LinePiece[]): number => -pt(pieces[1]!.baselineShiftPx ?? 0);

    it('raises a run by half a point of the LINE for each unit', async () => {
        // Twelve units came out at 6.90pt, not the 6.00 that half-points of
        // the point size would give — the line stands 1.15 times the size, and
        // the raise is measured in that.
        const lines = await drawn('run-position.docx');

        expect(riseOf(lines[1]!)).toBeCloseTo(6.90, 1);
    });

    it('lowers a run by the same distance the other way', async () => {
        expect(riseOf((await drawn('run-position.docx'))[2]!)).toBeCloseTo(-6.90, 1);
    });

    it('is NOT the superscript rise, which is its own smaller thing', async () => {
        // 3.95pt and a shrunken font, against 6.90 and no shrink at all.
        const lines = await drawn('run-position.docx');

        expect(riseOf(lines[3]!)).toBeCloseTo(3.95, 1);
        expect(lines[3]![1]!.sizePx).toBeLessThan(lines[3]![0]!.sizePx!);
    });

    it('WINS over w:vertAlign, size and all, where a run states both', async () => {
        // Measured: a run with both came out at the hand-set height and at
        // full size — its neighbour started exactly where the plain line's
        // did, which a shrunken run's would not.
        const lines = await drawn('run-position.docx');

        expect(riseOf(lines[4]!)).toBeCloseTo(6.90, 1);
        expect(lines[4]![1]!.sizePx).toBe(lines[4]![0]!.sizePx);
        // Full size means full width: the run after it starts where the
        // hand-raised line's does and not where the shrunken one's does.
        expect(lines[4]![1]!.widthPx).toBeCloseTo(lines[1]![1]!.widthPx, 6);
        expect(lines[4]![1]!.widthPx).toBeGreaterThan(lines[3]![1]!.widthPx);
    });

    it('scales with the units asked for and NOT with the font size', async () => {
        // 2, 6, 12 and 24 units came out at 1.15, 3.45, 6.90 and 13.85 — and
        // twelve units at twice the size came out at 6.85, the same raise.
        const lines = await drawn('run-position-scale.docx');

        expect(riseOf(lines[1]!)).toBeCloseTo(1.15, 1);
        expect(riseOf(lines[2]!)).toBeCloseTo(3.45, 1);
        expect(riseOf(lines[3]!)).toBeCloseTo(6.90, 1);
        expect(riseOf(lines[4]!)).toBeCloseTo(13.85, 0);
        expect(riseOf(lines[5]!)).toBeCloseTo(-3.45, 1);
        expect(riseOf(lines[6]!)).toBeCloseTo(6.85, 1);
    });

    it('takes the whole raise out of the LINE, which grows to hold it', async () => {
        // LibreOffice's baselines stepped 18.40pt where a raised run was
        // involved, against 11.50 for an ordinary pair — the whole 6.90 on top
        // of the line. Asserted as HEIGHTS rather than steps, because `yPx` is
        // a line's top: the gap between two tops is the height of the FIRST of
        // them, and reading it as the second's is off by one line.
        const opened = await openWordFile(file('run-position.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const heightOf = (index: number): number => pt(page!.lines[index]!.heightPx);

        expect(Math.abs(heightOf(0) - 11.50)).toBeLessThan(TOLERANCE_PT);
        expect(Math.abs(heightOf(1) - 18.40)).toBeLessThan(TOLERANCE_PT);
        // A lowered run grows the line by just as much, downward.
        expect(Math.abs(heightOf(2) - 18.40)).toBeLessThan(TOLERANCE_PT);
        // A superscript does not: it is smaller as well as raised, and fits
        // inside the line the plain runs beside it already ask for.
        expect(Math.abs(heightOf(3) - 11.50)).toBeLessThan(TOLERANCE_PT);
    });
});

describe('w:pgBorders, end to end', () => {
    // A three-point border (`w:sz="24"`) 24 points clear, on A4 with 567-twip
    // margins. LibreOffice's numbers are the CENTRES of the rules it drew.
    const TOLERANCE_PT = 0.12;
    const PAGE_HEIGHT_PT = 841.861;
    const pt = (px: number): number => px * 72 / 96;

    const borderOf = async (name: string): Promise<PlacedParagraphBorder> => {
        const opened = await openWordFile(file(name), FONTS);
        const [page] = layoutSections(opened.document.sections);

        return page!.pageBorder!;
    };

    /** Left, right, top and bottom in points, y counting from the page FOOT. */
    const edges = (box: PlacedParagraphBorder): number[] => [
        pt(box.leftPx),
        pt(box.rightPx),
        PAGE_HEIGHT_PT - pt(box.topPx),
        PAGE_HEIGHT_PT - pt(box.bottomPx),
    ];

    const near = (ours: readonly number[], theirs: readonly number[]): void => {
        ours.forEach((value, index) => {
            expect(Math.abs(value - theirs[index]!)).toBeLessThan(TOLERANCE_PT);
        });
    };

    it('measures a `page` border from the edge of the PAPER', async () => {
        // 24 points of clear paper, then the rule — whose centre is half a
        // width further in, at 25.5. LibreOffice drew 25.5, 569.75, 816.489
        // and 25.539.
        near(edges(await borderOf('page-border-from-page.docx')),
            [25.5, 569.75, 816.389, 25.539]);
    });

    it('measures a `text` border from the WRITING AREA, outward', async () => {
        // The same border against a 28.35pt margin: 24 points clear OUTSIDE
        // the text, so the rule centre lands at 2.85 — 22.6pt from where the
        // page-relative one went, which is why the default matters.
        near(edges(await borderOf('page-border-from-text.docx')),
            [2.9, 592.45, 838.989, 2.839]);
    });

    it('leaves the text exactly where a page with no border puts it', async () => {
        // LibreOffice printed the same three lines at 804.14, 792.64 and
        // 781.14 with the border, without it, and either way round.
        for (const name of ['page-border-from-page.docx', 'page-border-from-text.docx']) {
            const opened = await openWordFile(file(name), FONTS);
            const [page] = layoutSections(opened.document.sections);
            const first = page!.lines[0]!;

            expect(page!.lines).toHaveLength(3);
            // 841.861 less the 804.14 LibreOffice put the first baseline at.
            expect(Math.abs(pt(first.yPx + first.baselinePx) - 37.72))
                .toBeLessThan(TOLERANCE_PT);
        }
    });

    it('draws four rules, mitred at the corners like any other box', async () => {
        // The same drawing path a paragraph border goes through: LibreOffice
        // ran the top from 24.0 to 571.25 against sides at 25.5 and 569.75 —
        // half a width past each end, both ways.
        const opened = await openWordFile(file('page-border-from-page.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const rules = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind);

        expect(rules).toHaveLength(4);
        const horizontal = rules.filter((rule) => rule.y1Px === rule.y2Px);
        expect(horizontal).toHaveLength(2);
        for (const rule of horizontal) {
            expect(pt(Math.min(rule.x1Px, rule.x2Px))).toBeCloseTo(24.0, 1);
            expect(pt(Math.max(rule.x1Px, rule.x2Px))).toBeCloseTo(571.25, 1);
        }
    });
});

describe('w:lnNumType, end to end', () => {
    // A4, 567-twip margins, 10pt text. LibreOffice's x is where each number's
    // glyphs BEGIN, so a two-digit number starts further left than a one-digit
    // one — which is what says they are right-aligned.
    const TOLERANCE_PT = 0.12;
    const pt = (px: number): number => px * 72 / 96;

    const numbered = async (name: string): Promise<Page[]> => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections);
    };

    const textsOn = (page: Page): string[] => page.lineNumbers.map((n) => n.run.text);

    it('numbers every body line, an EMPTY paragraph among them', async () => {
        // LibreOffice printed 1 to 5 down a page holding four paragraphs, one
        // of them empty, and a table. The empty one took a number.
        const [page] = await numbered('line-numbers.docx');

        expect(textsOn(page!)).toEqual(['1', '2', '3', '4', '5']);
    });

    it('does NOT number a line inside a table', async () => {
        // Six lines of text are drawn on this page and only five numbered:
        // LibreOffice passed the cell's line over silently, and the number it
        // would have taken went to the paragraph below instead.
        const [page] = await numbered('line-numbers.docx');
        const celled = page!.rows.flatMap((row) => row.cells.flatMap((cell) => cell.lines));

        expect(celled).toHaveLength(1);
        expect(page!.lineNumbers).toHaveLength(5);
        // No number stands on the cell's own baseline.
        const cellBaseline = celled[0]!.yPx + celled[0]!.baselinePx;
        expect(page!.lineNumbers.map((n) => n.baselinePx)).not.toContain(cellBaseline);
    });

    it('puts each number’s RIGHT edge the stated distance in from the text', async () => {
        // 360 twips is 18pt, and the writing area starts at 28.35 — so the
        // right edge lands at 10.35, where LibreOffice drew 10.45.
        const [page] = await numbered('line-numbers.docx');
        const first = page!.lineNumbers[0]!;
        const width = first.run.font.measureAdvance(first.run.text, first.run.sizePx).widthPt;

        expect(Math.abs(pt(first.xPx + width) - 10.35)).toBeLessThan(0.2);
    });

    it('right-aligns them, so a two-digit number reaches further left', async () => {
        // 6 and 8 began at 4.90 and 10 at −0.65: the same right edge, a wider
        // number. Left-align them and all three would start together.
        const [page] = await numbered('line-numbers-every-two.docx');
        const [six, eight, ten] = page!.lineNumbers;

        expect(textsOn(page!)).toEqual(['6', '8', '10']);
        expect(six!.xPx).toBe(eight!.xPx);
        expect(ten!.xPx).toBeLessThan(six!.xPx);
    });

    it('adds w:start to the count rather than starting at it', async () => {
        // `start=5` with `countBy=2` printed 6, 8, 10 against lines one, three
        // and five — so the first line is start + 1, and a number is printed
        // where the count divides by countBy.
        const [page] = await numbered('line-numbers-every-two.docx');
        const baselines = page!.lineNumbers.map((n) => n.baselinePx);
        const lines = page!.lines.map((line) => line.yPx + line.baselinePx);

        expect(baselines).toEqual([lines[0], lines[2], lines[4]]);
    });

    it('leaves half a centimetre where no distance is stated', async () => {
        // Measured at 14.05pt in from the writing area, not the 18 a quarter
        // of an inch would have given.
        const [page] = await numbered('line-numbers-default-gap.docx');
        const first = page!.lineNumbers[0]!;
        const width = first.run.font.measureAdvance(first.run.text, first.run.sizePx).widthPt;

        expect(Math.abs(pt(first.xPx + width) - 14.30)).toBeLessThan(0.2);
    });

    it('starts again on each page when nothing says otherwise', async () => {
        // The default is `newPage`, measured: page two began again at 1.
        const pages = await numbered('line-numbers-per-page.docx');

        expect(pages.length).toBeGreaterThan(1);
        expect(textsOn(pages[1]!)[0]).toBe('1');
    });

    it('runs the count on across pages where it says continuous', async () => {
        // Page two picked up at 69, which is where page one left off.
        const pages = await numbered('line-numbers-continuous.docx');

        expect(textsOn(pages[1]!)[0]).toBe('69');
        expect(textsOn(pages[0]!)).toHaveLength(68);
    });

    it('draws each number in the font of the line it stands beside', async () => {
        const [page] = await numbered('line-numbers.docx');
        const number = page!.lineNumbers[0]!;
        const piece = page!.lines[0]!.line.pieces[0]!;

        expect(number.run.sizePx).toBe(piece.sizePx);
        expect(number.run.font).toBe(piece.font);
    });

    it('numbers nothing at all where the section asks for none', async () => {
        const [page] = await numbered('baseline.docx');

        expect(page!.lineNumbers).toEqual([]);
    });

    it('DRAWS them, in the margin, left of everything else on the page', async () => {
        // Every test above reads the model. This one renders: a number placed
        // and never drawn would satisfy all of them and print nothing.
        const [page] = await numbered('line-numbers.docx');
        const ops = renderPage(page!).ops.filter((op): op is TextOp => 'text' === op.kind);
        const digits = ops.filter((op) => /^\d+$/u.test(op.text));

        expect(digits.map((op) => op.text)).toEqual(['1', '2', '3', '4', '5']);
        // In the margin: every one of them left of the writing area at 28.35.
        for (const digit of digits) {
            expect(pt(digit.xPx)).toBeLessThan(28.35);
        }
        // And on the baselines the model put them at.
        expect(digits.map((op) => op.yPx)).toEqual(page!.lineNumbers.map((n) => n.baselinePx));
    });
});

describe('w:bdr, drawn end to end', async () => {
    // A4, 567-twip margins, 10pt Liberation Serif. Each line holds a plain
    // run, a BOXED run, and a plain run — so the box's edges and the advance
    // it costs are both on the page beside a control.
    const TOLERANCE_PT = 0.15;
    const pt = (px: number): number => px * 72 / 96;

    const laid = async (name: string): Promise<Page[]> => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections);
    };

    /** The vertical rules of each box on the page, left to right, in points. */
    const boxesOf = (page: Page): { left: number; right: number; top: number }[] => {
        const rules = renderPage(page).ops.filter((op): op is LineOp => 'line' === op.kind);
        const boxes: { left: number; right: number; top: number }[] = [];

        for (let index = 0; index + 3 < rules.length; index += 4) {
            const [top, , left, right] = rules.slice(index, index + 4);

            boxes.push({
                left: pt(left!.x1Px),
                right: pt(right!.x1Px),
                top: pt(top!.y1Px),
            });
        }

        return boxes;
    };

    it('draws one box per boxed run, and none round the plain ones', async () => {
        // Four boxed runs on the page, five plain lines between them.
        const [page] = await laid('run-border.docx');

        expect(boxesOf(page!)).toHaveLength(4);
    });

    it('stands the rules half a width outside the text, on both sides', async () => {
        // LibreOffice's space-nought box: a 19.25pt span for a run whose own
        // advance is 18.30, which is half a one-point rule at each end.
        const [page] = await laid('run-border.docx');
        const [first] = boxesOf(page!);

        expect(first!.right - first!.left).toBeCloseTo(19.25, 0);
    });

    it('grows the box a point for every point of space, on each side', async () => {
        // LibreOffice's four boxes measured 19.25, 23.25, 27.25 and 25.25.
        // 19.25, 23.25, 27.25 for spaces of 0, 2 and 4 — twice the space, and
        // the same 25.25 for a three-point rule at space 2 as the arithmetic
        // says it must be.
        const [page] = await laid('run-border.docx');
        const widths = boxesOf(page!).map((box) => box.right - box.left);

        expect(widths.map((w) => Math.round(w * 100) / 100))
            .toEqual([19.25, 23.25, 27.25, 25.25].map((w) => expect.closeTo(w, 0)));
    });

    it('makes the line taller by the room it keeps, above and below', async () => {
        // LibreOffice's lines: 11.5pt plain, then 13.5, 17.5 and 21.5.
        // 11.5pt lines became 13.5, 17.5 and 21.5 for room of 1, 3 and 5.
        const [page] = await laid('run-border.docx');
        const heights = page!.lines.map((line) => pt(line.heightPx));

        expect(heights[0]).toBeCloseTo(11.5, 1);
        expect(heights[2]).toBeCloseTo(13.5, 1);
        expect(heights[4]).toBeCloseTo(17.5, 1);
        expect(heights[6]).toBeCloseTo(21.5, 1);
        expect(heights[8]).toBeCloseTo(21.5, 1);
    });

    it('costs the run beside it the room, so the next run clears the box', async () => {
        // LibreOffice's advances: 19.30, 21.30, 23.30 and 23.30 again.
        // The advance from the boxed run's text to the run after it grew by
        // exactly the room: 19.30, 21.30, 23.30 for one-point rules at spaces
        // 0, 2 and 4, and 23.30 again for a three-point rule at space 2.
        const [page] = await laid('run-border.docx');
        const advance = (lineIndex: number): number => {
            const pieces = page!.lines[lineIndex]!.line.pieces;

            return pt(pieces[1]!.widthPx + (pieces[1]!.borderRoomPx ?? 0));
        };

        expect(advance(2)).toBeCloseTo(19.30, 0);
        expect(advance(4)).toBeCloseTo(21.30, 0);
        expect(advance(6)).toBeCloseTo(23.30, 0);
        expect(advance(8)).toBeCloseTo(23.30, 0);
    });

    it('keeps its own size beside a MUCH bigger run on the same line', async () => {
        // The defect this closes: the box used to hug the LINE, so a 10pt
        // boxed run beside a 20pt one drew a box grown to the 20pt line.
        // LibreOffice drew 18.45pt either way, and in the same place relative
        // to the boxed run's OWN baseline — 12.9 above it and 5.55 below.
        const [page] = await laid('run-border-tall-neighbour.docx');
        const rules = renderPage(page!).ops.filter((op): op is LineOp => 'line' === op.kind);
        const boxes: { top: number; bottom: number }[] = [];

        for (let index = 0; index + 3 < rules.length; index += 4) {
            const [top, bottom] = rules.slice(index, index + 4);
            boxes.push({ top: pt(top!.y1Px), bottom: pt(bottom!.y1Px) });
        }

        // Two boxed runs on the page: one beside a 20pt run, one beside a 10pt.
        expect(boxes).toHaveLength(2);
        expect(boxes[0]!.bottom - boxes[0]!.top).toBeCloseTo(18.45, 0);
        expect(boxes[1]!.bottom - boxes[1]!.top).toBeCloseTo(18.45, 0);

        // And each sits the same distance above its own line's boxed text.
        const lines = page!.lines.filter(
            (line) => line.line.pieces.some((piece) => undefined !== piece.border));

        expect(lines).toHaveLength(2);
        lines.forEach((line, index) => {
            const baseline = pt(line.yPx + line.baselinePx);

            expect(baseline - boxes[index]!.top).toBeCloseTo(12.9, 0);
            expect(boxes[index]!.bottom - baseline).toBeCloseTo(5.55, 0);
        });
    });

    it('draws a COMPLETE box on every line a boxed run breaks over', async () => {
        // Three lines, three closed boxes — not one outline left open at the
        // end of a line and picked up on the next.
        const [page] = await laid('run-border-wrapped.docx');

        expect(page!.lines.length).toBeGreaterThan(3);
        expect(boxesOf(page!)).toHaveLength(3);
    });

    it('stands the rules above and below the TEXT, not outside the room', async () => {
        // LibreOffice's boxes measured 12.45, 16.45, 20.45 and 18.45pt tall
        // for rooms of 1, 3, 5 and 5 — the line box plus a standoff at each
        // end, NOT the whole line the room had already grown.
        const [page] = await laid('run-border.docx');
        const rules = renderPage(page!).ops.filter((op): op is LineOp => 'line' === op.kind);
        const heights: number[] = [];

        for (let index = 0; index + 3 < rules.length; index += 4) {
            const [top, bottom] = rules.slice(index, index + 4);
            heights.push(pt(bottom!.y1Px - top!.y1Px));
        }

        expect(heights[0]).toBeCloseTo(12.45, 0);
        expect(heights[1]).toBeCloseTo(16.45, 0);
        expect(heights[2]).toBeCloseTo(20.45, 0);
        expect(heights[3]).toBeCloseTo(18.45, 0);
    });

    it('runs each rule half the crossing rule past the corner', async () => {
        // LibreOffice's space-nought box ran its top from 35.600 to 55.850
        // against sides at 36.100 and 55.350 — half a width past, both ends.
        const [page] = await laid('run-border.docx');
        const [top, , left, right] = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind).slice(0, 4);

        expect(pt(Math.min(top!.x1Px, top!.x2Px))).toBeCloseTo(pt(left!.x1Px) - 0.5, 1);
        expect(pt(Math.max(top!.x1Px, top!.x2Px))).toBeCloseTo(pt(right!.x1Px) + 0.5, 1);
    });

    it('boxes the whole SPAN on each line, not each word of it', async () => {
        // Eleven words to a line, one box round the lot: LibreOffice's first
        // ran 36.5 to 548.95, and its continuations 29.9 to 542.35 and 425.65.
        // A box per word would stop at the first of them.
        const [page] = await laid('run-border-wrapped.docx');
        const boxes = boxesOf(page!);

        expect(boxes.map((box) => Math.round(box.left)))
            .toEqual([37, 30, 30].map((value) => expect.closeTo(value, 0)));
        expect(boxes.map((box) => Math.round(box.right)))
            .toEqual([549, 542, 426].map((value) => expect.closeTo(value, 0)));
    });

    it('indents a continuation line to leave room for its own left edge', async () => {
        // LibreOffice began lines two and three at 33.35 against a 28.35
        // margin: the space and the whole rule, five points of it.
        const [page] = await laid('run-border-wrapped.docx');
        const boxes = boxesOf(page!);
        const ops = renderPage(page!).ops.filter((op): op is TextOp => 'text' === op.kind);

        // By LINE, not by op: the run is cut into a piece per word, so eleven
        // of them stand on the first line alone and "the second `wwwww`" is
        // the second WORD, not the second line.
        const baselines = [...new Set(ops.map((op) => op.yPx))].sort((a, b) => a - b);
        const leftmostOn = (baseline: number): number => Math.min(
            ...ops.filter((op) => op.yPx === baseline).map((op) => op.xPx));

        // Line 0 is the plain 'A'; line 1 opens with the plain 'B' run and so
        // starts ON the margin. Lines 2 and 3 are the run's continuations, and
        // those are the ones that have to leave room for their own left edge.
        expect(pt(leftmostOn(baselines[1]!))).toBeCloseTo(28.35, 0);
        expect(pt(leftmostOn(baselines[2]!))).toBeCloseTo(33.35, 0);
        expect(pt(leftmostOn(baselines[3]!))).toBeCloseTo(33.35, 0);
        expect(boxes[1]!.left).toBeCloseTo(29.9, 0);
    });
});

describe('what the tidy fixtures could not see', () => {
    // Both of these passed every test in the suite that shipped them, because
    // both fixtures stated the value at which the rule makes no difference.
    const TOLERANCE_PT = 0.15;
    const PAGE_HEIGHT_PT = 841.861;
    const pt = (px: number): number => px * 72 / 96;

    const pageOf = async (name: string): Promise<Page> => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections)[0]!;
    };

    it('keeps a bordered paragraph’s bottom rule INSIDE the writing area', async () => {
        // The boundary, built to sit on it: 11.5pt lines in a 785.16pt column,
        // and a box needing its line plus 13pt of rule and space either side —
        // 37.5 in all. LibreOffice kept it on a page holding 65 filler lines
        // and moved it off one holding 66. Reserve only the TOP room, as this
        // did, and 66 keeps it: the rule then prints below the foot.
        const after = async (fillers: number): Promise<Page[]> => {
            const opened = await openWordFile(
                file(`border-at-page-foot-${fillers}.docx`), FONTS);

            return layoutSections(opened.document.sections);
        };

        expect(await after(65)).toHaveLength(1);
        expect(await after(66)).toHaveLength(2);
    });

    it('leaves the rule of the one that FITS clear of the foot', async () => {
        // LibreOffice printed its bottom rule with an outer edge at 29.539
        // against a writing area ending at 28.35 — inside it, with a point to
        // spare, which is the whole of what the reserve buys.
        const opened = await openWordFile(file('border-at-page-foot-65.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.paragraphBorders;
        const footPx = page!.geometry.heightPx - page!.geometry.marginBottomPx;
        const outerPx = box!.bottomPx + (box!.borders.bottom?.widthPx ?? 0) / 2;

        expect(outerPx).toBeLessThanOrEqual(footPx);
    });

    it('splits a box into two COMPLETE boxes, and reserves for both', async () => {
        // The fixture is built to sit on the boundary: 64 filler lines, then a
        // paragraph boxed with nothing above it and 25pt of rule and space
        // below. LibreOffice took TWO of its lines onto page one — four would
        // have fitted with no reserve at all — and drew four rules on each
        // page, top and bottom either side of the break.
        const opened = await openWordFile(file('border-split-at-foot.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages).toHaveLength(2);

        const boxedOn = (page: Page): number => page.lines.filter(
            (line) => 'w'.repeat(10) === line.line.pieces[0]?.text.slice(0, 10)).length;

        expect(boxedOn(pages[0]!)).toBe(2);

        // A whole box on each page, and each one drawn as four rules.
        for (const page of pages) {
            const [box] = page.paragraphBorders;

            expect([box!.opensHere, box!.closesHere]).toEqual([true, true]);
            expect(renderPage(page).ops.filter((op) => 'line' === op.kind)).toHaveLength(4);
        }
    });

    it('keeps the split box’s own bottom rule inside the page it ends', async () => {
        // Page one's bottom rule printed with its outer edge at 28.539 against
        // a writing area ending at 28.35 — which is what the reserve buys, and
        // what taking it only for the paragraph's LAST line would have missed.
        const opened = await openWordFile(file('border-split-at-foot.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.paragraphBorders;
        const footPx = page!.geometry.heightPx - page!.geometry.marginBottomPx;

        expect(box!.bottomPx + (box!.borders.bottom?.widthPx ?? 0) / 2)
            .toBeLessThanOrEqual(footPx);
    });

    it('draws a bordered paragraph INSIDE a cell, which used to draw nothing', async () => {
        // The gap the code itself gave away: paragraph boxes were collected by
        // the page flow, and a cell's content goes through `stackBlocks`, which
        // collected none. LibreOffice drew the box at 33.1 to 172.55 — half a
        // one-point rule outside a text column running 33.65 to 172.15.
        const opened = await openWordFile(file('cell-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [cell] = page!.rows[0]!.cells;

        expect(cell!.paragraphBorders).toHaveLength(1);

        const [box] = cell!.paragraphBorders;
        expect(pt(box!.leftPx)).toBeCloseTo(33.1, 0);
        expect(pt(box!.rightPx)).toBeCloseTo(172.55, 0);
        expect(pt(box!.bottomPx - box!.topPx)).toBeCloseTo(12.45, 0);
    });

    it('makes the ROW taller by the room the box keeps', async () => {
        // LibreOffice's row measured 14.5pt between the CENTRES of its own
        // rules, where a plain one measures 12.5. Those centres sit half a
        // table rule outside the row's content either end, so the content is
        // 13.5 against a plain 11.5 — the boxed paragraph's one point of rule
        // above and one below. Comparing a content height to a rule-to-rule
        // distance is how this first read as a point short.
        const opened = await openWordFile(file('cell-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        expect(pt(page!.rows[0]!.heightPx)).toBeCloseTo(14.5 - 1, 0);
    });

    it('DRAWS it, rather than merely placing it', async () => {
        const opened = await openWordFile(file('cell-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.rows[0]!.cells[0]!.paragraphBorders;
        const rules = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind);

        // Four of the box's own among the table's: two verticals standing on
        // the box's own edges, which no table rule does.
        expect(rules.filter((rule) => rule.x1Px === rule.x2Px
            && Math.abs(rule.x1Px - box!.leftPx) < 0.01)).toHaveLength(1);
        expect(rules.filter((rule) => rule.x1Px === rule.x2Px
            && Math.abs(rule.x1Px - box!.rightPx) < 0.01)).toHaveLength(1);
    });

    it('leaves a cell of plain paragraphs with no box at all', async () => {
        const opened = await openWordFile(file('cell-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);

        expect(page!.rows[0]!.cells[1]!.paragraphBorders).toEqual([]);
    });

    it('runs ONE box round a cell’s bordered paragraphs, wraps and all', async () => {
        // Two bordered paragraphs in a cell, the second wrapping onto a third
        // line. LibreOffice drew a single outline 37.45pt tall — three 11.5pt
        // lines and a 1.5 standoff either end — with NO rule between the two
        // paragraphs, exactly as it does on the page. A three-point rule, so
        // the standoff is wider than the tolerance that let the last fixture
        // pass a box drawn without one.
        const opened = await openWordFile(file('cell-paragraph-border-run.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [cell] = page!.rows[0]!.cells;

        expect(cell!.paragraphBorders).toHaveLength(1);
        expect(cell!.lines).toHaveLength(3);

        const [box] = cell!.paragraphBorders;
        expect(pt(box!.bottomPx - box!.topPx)).toBeCloseTo(37.45, 0);
        expect(box!.innerYPx).toEqual([]);
    });

    it('stands a cell’s box off its text on BOTH sides', async () => {
        // LibreOffice drew the sides at 32.1 and 173.55.
        // 32.1 and 173.55 against a text column running 33.65 to 172.15 — a
        // point and a half outside it either way, which is half the rule.
        const opened = await openWordFile(file('cell-paragraph-border-run.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.rows[0]!.cells[0]!.paragraphBorders;

        expect(pt(box!.leftPx)).toBeCloseTo(32.1, 0);
        expect(pt(box!.rightPx)).toBeCloseTo(173.55, 0);
    });

    it('draws a bordered paragraph in a HEADER, which used to draw nothing', async () => {
        // The same blind spot as a cell's, one level over: a header is stacked
        // by the code a cell is stacked by, and the boxes it gathered were
        // dropped where the furniture was placed. LibreOffice drew this one
        // 14.45pt tall — an 11.5pt line and a 1.5 standoff either end for a
        // three-point rule — with its sides at 70.5 and 524.75.
        const opened = await openWordFile(file('header-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const header = page!.header!;

        expect(header.paragraphBorders).toHaveLength(1);

        const [box] = header.paragraphBorders;
        expect(pt(box!.leftPx)).toBeCloseTo(70.5, 0);
        expect(pt(box!.rightPx)).toBeCloseTo(524.75, 0);
        expect(pt(box!.bottomPx - box!.topPx)).toBeCloseTo(14.45, 0);
    });

    it('puts the header’s box on the HEADER, not at the top of the body', async () => {
        // In page coordinates: the box has to have moved with the furniture.
        const opened = await openWordFile(file('header-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const header = page!.header!;
        const [box] = header.paragraphBorders;

        expect(box!.topPx).toBeGreaterThanOrEqual(header.topPx - 5);
        expect(box!.bottomPx).toBeLessThanOrEqual(page!.lines[0]!.yPx);
    });

    it('DRAWS it, rather than merely placing it', async () => {
        const opened = await openWordFile(file('header-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.header!.paragraphBorders;
        const rules = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind);

        expect(rules.filter((rule) => rule.x1Px === rule.x2Px
            && Math.abs(rule.x1Px - box!.leftPx) < 0.01)).toHaveLength(1);
        expect(rules.filter((rule) => rule.y1Px === rule.y2Px
            && Math.abs(rule.y1Px - box!.topPx) < 0.01)).toHaveLength(1);
    });

    it('ignores a paragraph border inside a FOOTNOTE, as LibreOffice does', async () => {
        // Probed rather than reasoned about, and the answer went the other way
        // from the header and the cell: LibreOffice drew ONE rule on the page
        // and it was the footnote separator. No box, and no room taken either
        // — the note's text printed at exactly the place it does without the
        // border. So this is not a gap; building it would make us diverge.
        //
        // Asserted because the absence is now load-bearing: notes are stacked
        // by `stackBlocks`, which gathers boxes for a cell and a header, and
        // anything that later unified the three would start drawing these
        // without a word.
        const opened = await openWordFile(file('footnote-paragraph-border.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const rules = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind);

        // The border IS in the file and IS read — without this the test would
        // pass just as well on a fixture that never had one.
        const declared = [...opened.document.footnotes.values()][0]![0]!;
        expect(isTable(declared) ? undefined : declared.style.borders?.top).toBeDefined();

        // The separator, and nothing else.
        expect(rules).toHaveLength(1);
        expect(rules[0]!.x2Px - rules[0]!.x1Px).toBe(page!.footnotes!.separatorWidthPx);

        // And the note sits where an unbordered one sits: 30.45 up from the
        // foot of a 226.8pt page, which is what the unbordered fixture measured.
        const note = page!.footnotes!.lines[0]!;
        expect(226.8 - pt(note.yPx + note.baselinePx)).toBeCloseTo(30.45, 1);
    });

    it('rules between a CELL’s bordered paragraphs, as it does on the page', async () => {
        // The hole this fills: the cell's tracker set `innerYPx` to
        // an empty list and never added to it, so `w:between` inside a cell
        // drew nothing. LibreOffice draws it at 786.989 inside a box running
        // 800.489 to 773.539, and steps the pair 14.5pt where an unruled pair
        // steps 11.5 — the three-point rule's whole width out of the flow.
        const opened = await openWordFile(
            file('cell-paragraph-border-between.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [cell] = page!.rows[0]!.cells;
        const [box] = cell!.paragraphBorders;
        const first = cell!.lines[0]!;

        expect(cell!.paragraphBorders).toHaveLength(1);
        expect(box!.innerYPx).toHaveLength(1);

        // Asserted against the cell's OWN lines rather than the page, because
        // a cell's content sits about a point higher here than LibreOffice
        // puts it — a separate defect, recorded on its own, which an absolute
        // assertion would fold into this one and hide.
        expect(pt(box!.innerYPx[0]! - (first.yPx + first.heightPx)))
            .toBeCloseTo(1.5, 1);
        expect(pt(box!.topPx)).toBeCloseTo(pt(first.yPx) - 0.5, 1);

        // And the rule is out of the flow, not drawn over the text: the pair
        // steps 14.5pt where an unruled pair steps 11.5.
        expect(pt(cell!.lines[1]!.yPx - first.yPx)).toBeCloseTo(14.5, 1);
    });

    it('DRAWS the cell’s between rule, at the width it asked for', async () => {
        const opened = await openWordFile(
            file('cell-paragraph-border-between.docx'), FONTS);
        const [page] = layoutSections(opened.document.sections);
        const [box] = page!.rows[0]!.cells[0]!.paragraphBorders;
        const rules = renderPage(page!).ops
            .filter((op): op is LineOp => 'line' === op.kind);

        // Three points wide, where the box's own sides are one.
        expect(rules.filter((rule) => rule.y1Px === rule.y2Px
            && Math.abs(rule.y1Px - box!.innerYPx[0]!) < 0.01
            && Math.abs(pt(rule.widthPx) - 3) < 0.01)).toHaveLength(1);
    });

    it('puts a SPACED w:between half a rule below the text above it', async () => {
        // With six points of space and a three-point rule, LibreOffice still
        // drew the rule at 787.989 — half a width below the upper line — and
        // put the whole six points BELOW it, stepping the pair 20.5pt apart
        // where an unspaced pair steps 14.5. The earlier fixture stated no
        // space at all, where `width / 2` and `space + width / 2` agree.
        const page = await pageOf('paragraph-border-between-spaced.docx');
        const [box] = page.paragraphBorders;

        expect(box!.innerYPx).toHaveLength(1);
        expect(Math.abs(PAGE_HEIGHT_PT - pt(box!.innerYPx[0]!) - 787.989))
            .toBeLessThan(TOLERANCE_PT);

        // And the space is out of the flow all the same: 11.5 + 6 + 3.
        const step = pt(page.lines[2]!.yPx - page.lines[1]!.yPx);
        expect(Math.abs(step - 20.5)).toBeLessThan(TOLERANCE_PT);
    });

    it('measures w:tblInd against the LEADING CELL’s own margin', async () => {
        // The first cell overrides its left margin to 500 twips. LibreOffice
        // put the table's border at 39.4 — margin plus indent less the CELL's
        // 25pt — where the table's own 5.75 would have given 58.6. The text
        // still lands on margin plus indent either way, which is why a fixture
        // without the override cannot tell the two apart.
        const page = await pageOf('table-indent-cell-margin.docx');
        const [row] = page.rows;

        expect(pt(row!.cells[0]!.xPx)).toBeCloseTo(39.4, 0);
        expect(pt(row!.cells[0]!.lines[0]!.xPx)).toBeCloseTo(64.45, 0);
    });
});

describe('the room a table keeps for its rules', () => {
    // A line, a three-row table of one-point rules, a line. Every number is
    // LibreOffice's, off the same A4 page the rest of these use.
    const PAGE_HEIGHT_PT = 841.861;
    const TOLERANCE_PT = 0.12;
    const at = (px: number): number => PAGE_HEIGHT_PT - px * 72 / 96;

    const laid = async (): Promise<Page> => {
        const opened = await openWordFile(file('table-rule-gaps.docx'), FONTS);

        return layoutSections(opened.document.sections)[0]!;
    };

    it('keeps a whole rule-width between the text above and the first row', async () => {
        // The line above ends at 801.97 and the first row's text begins at
        // 800.97 — a gap of one point for a one-point rule, which is drawn
        // down the middle of it at 801.489.
        const page = await laid();
        const above = page.lines[0]!;
        const first = page.rows[0]!;

        expect(at(above.yPx + above.heightPx)).toBeCloseTo(801.97, 1);
        expect(at(first.yPx)).toBeCloseTo(800.97, 1);
    });

    it('keeps one between every pair of rows as well, not just at the ends', async () => {
        // Row steps of 12.5pt against an 11.5pt line: the inside rules take
        // their width out of the flow exactly as the outer ones do.
        const page = await laid();
        const tops = page.rows.map((row) => at(row.yPx));

        expect(tops[0]! - tops[1]!).toBeCloseTo(12.5, 1);
        expect(tops[1]! - tops[2]!).toBeCloseTo(12.5, 1);
    });

    it('leaves the line AFTER the table clear of its bottom rule', async () => {
        // The last row's text ends at 764.48 and the line below begins at
        // 763.47, with the rule at 763.989 between them.
        const page = await laid();
        const last = page.rows[2]!;
        const below = page.lines[1]!;

        expect(at(last.yPx + last.heightPx)).toBeCloseTo(764.48, 1);
        expect(at(below.yPx)).toBeCloseTo(763.47, 1);
    });

    it('adds up: 38.5pt for three lines and four rules', async () => {
        // The whole table, from the bottom of the line above to the top of the
        // line below — 3 x 11.5 plus 4 x 1, which is what says every rule has
        // a gap and no rule has two.
        const page = await laid();
        const above = page.lines[0]!;
        const below = page.lines[1]!;
        const span = at(above.yPx + above.heightPx) - at(below.yPx);

        expect(Math.abs(span - 38.5)).toBeLessThan(TOLERANCE_PT);
    });

    it('gives each gap the width of the rule that SITS in it', async () => {
        // The fixture above states one width for every side, so `top`,
        // `insideH` and `bottom` are interchangeable in it and three mutants
        // survived. This one states 6pt on top, 1 inside and 3 below, and
        // LibreOffice left gaps of exactly 6.0, 1.0 and 3.0.
        const opened = await openWordFile(file('table-rule-gaps-mixed.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const [above, below] = page.lines;
        const rows = page.rows;

        // Above the first row: the TABLE's top, six points of it.
        expect(at(above!.yPx + above!.heightPx) - at(rows[0]!.yPx)).toBeCloseTo(6, 1);
        // Between rows: `insideH`, one point.
        expect(at(rows[0]!.yPx + rows[0]!.heightPx) - at(rows[1]!.yPx)).toBeCloseTo(1, 1);
        // Below the last: the bottom, three.
        expect(at(rows[2]!.yPx + rows[2]!.heightPx) - at(below!.yPx)).toBeCloseTo(3, 1);
    });

    it('takes the gap from a CELL’s own border where the table declares none', async () => {
        // No `w:tblBorders` at all, and one cell asking for a six-point top.
        // LibreOffice left six points — and moved the cell BESIDE it down with
        // it, because the gap belongs to the row and not to the cell.
        const opened = await openWordFile(file('table-rule-gap-from-cell.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const [above] = page.lines;
        const [row] = page.rows;

        expect(at(above!.yPx + above!.heightPx) - at(row!.yPx)).toBeCloseTo(6, 1);
        expect(row!.cells[1]!.lines[0]!.yPx).toBe(row!.cells[0]!.lines[0]!.yPx);
    });

    it('grows a MERGED cell’s box by the rules its span swallows', async () => {
        // A cell merged over three 11.5pt rows: LibreOffice drew its box from
        // 801.489 to 763.989, which is 37.5pt — 34.5 of text, the TWO
        // one-point inside rules the merge crosses, and half a rule outside at
        // either end. Sum the rows alone and the box stops 2pt above its foot.
        const opened = await openWordFile(file('vmerge-rule-gaps.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const merged = page.rows[0]!.cells[0]!;
        const last = page.rows[2]!;

        expect(merged.heightPx * 72 / 96).toBeCloseTo(36.5, 1);
        // Its foot is the last row's foot, gaps and all.
        expect(page.rows[0]!.yPx + merged.heightPx)
            .toBeCloseTo(last.yPx + last.heightPx, 6);
    });

    it('stops the inside rules AT the merged cell, not through it', async () => {
        // LibreOffice ran them from 227.4, the merged cell's right edge, and
        // drew the merged cell's own sides down the whole span.
        const opened = await openWordFile(file('vmerge-rule-gaps.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const merged = page.rows[0]!.cells[0]!;
        const inside = renderPage(page).ops
            .filter((op): op is LineOp => 'line' === op.kind)
            .filter((op) => op.y1Px === op.y2Px)
            .filter((op) => op.y1Px > page.rows[0]!.yPx
                && op.y1Px < page.rows[2]!.yPx + page.rows[2]!.heightPx);

        // TWO rules, by position — each is emitted twice, once as the upper
        // row's bottom and once as the lower row's top. That duplication is
        // older than this slice and was invisible while rows abutted; it is
        // recorded on its own rather than papered over here.
        const distinct = [...new Set(inside.map((rule) => Math.round(rule.y1Px * 100)))];

        expect(distinct).toHaveLength(2);
        for (const rule of inside) {
            expect(Math.min(rule.x1Px, rule.x2Px))
                .toBeGreaterThanOrEqual(merged.xPx + merged.widthPx - 0.01);
        }
    });

    it('draws ONE rule per shared edge — a row per rule, plus a top', async () => {
        // LibreOffice printed 63 horizontal rules for the 62 rows on the first
        // page of this table and 9 for the 8 on the second: one per row, and
        // one more for the top of each part. Drawing each shared edge twice —
        // once as the upper row's bottom and once as the lower row's top —
        // gives twice that, which is what this used to do.
        const opened = await openWordFile(file('table-split-rules.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const horizontals = (page: Page): number => renderPage(page).ops
            .filter((op): op is LineOp => 'line' === op.kind)
            .filter((op) => op.y1Px === op.y2Px).length;

        expect(pages).toHaveLength(2);
        expect(horizontals(pages[0]!)).toBe(pages[0]!.rows.length + 1);
        expect(horizontals(pages[1]!)).toBe(pages[1]!.rows.length + 1);
    });

    it('gives the half overleaf a top of its OWN', async () => {
        // A table split across pages is closed on both, the way a paragraph's
        // box is — the second page opens with a rule rather than
        // carrying on from the first.
        const opened = await openWordFile(file('table-split-rules.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const second = pages[1]!;
        const first = second.rows[0]!;
        const tops = renderPage(second).ops
            .filter((op): op is LineOp => 'line' === op.kind)
            .filter((op) => op.y1Px === op.y2Px && op.y1Px < first.yPx);

        expect(tops).toHaveLength(1);
    });

    it('draws each rule DOWN THE MIDDLE of the gap it kept', async () => {
        // 801.489, 788.989, 776.489 and 763.989 — each half a point outside
        // the content either side of it.
        const page = await laid();
        const rules = [...new Set(renderPage(page).ops
            .filter((op): op is LineOp => 'line' === op.kind)
            .filter((op) => op.y1Px === op.y2Px)
            .map((op) => Math.round(at(op.y1Px) * 1000) / 1000))].sort((a, b) => b - a);

        expect(rules).toHaveLength(4);
        [801.489, 788.989, 776.489, 763.989].forEach((expected, index) => {
            expect(Math.abs(rules[index]! - expected)).toBeLessThan(TOLERANCE_PT);
        });
    });
});

describe('w:docGrid, end to end', () => {
    // A section's typesetting grid. East Asian in origin and it moves LATIN
    // text: LibreOffice stepped 18pt for a 360-twip pitch where the same four
    // paragraphs ungridded step 11.5, and put the first baseline at 800.89
    // against 804.14 — 3.25 lower, which is half the 6.5 the grid added.
    const PAGE_HEIGHT_PT = 841.861;
    const TOLERANCE_PT = 0.12;
    const at = (px: number): number => PAGE_HEIGHT_PT - px * 72 / 96;

    const baselines = async (name: string): Promise<number[]> => {
        const opened = await openWordFile(file(name), FONTS);
        const [page] = layoutSections(opened.document.sections);

        return page!.lines.map((placed) => at(placed.yPx + placed.baselinePx));
    };

    it('steps a line of the grid’s pitch, not the font’s', async () => {
        const gridded = await baselines('doc-grid-lines.docx');
        const steps = gridded.slice(1).map((value, index) => gridded[index]! - value);

        expect(steps.map((step) => Math.round(step * 100) / 100)).toEqual([18, 18, 18]);
    });

    it('leaves an ungridded section on its font’s own line', async () => {
        const plain = await baselines('doc-grid-none.docx');
        const steps = plain.slice(1).map((value, index) => plain[index]! - value);

        expect(steps.map((step) => Math.round(step * 100) / 100)).toEqual([11.5, 11.5, 11.5]);
    });

    it('CENTRES the text in its grid line, rather than hanging it from the top', async () => {
        // The rule the pitch alone does not give: 800.89 against 804.14 is
        // 3.25 down, which is half of 18 less 11.5. Put the leading all above
        // or all below and the first baseline lands 3.25pt wrong on every page.
        const [firstGridded] = await baselines('doc-grid-lines.docx');
        const [firstPlain] = await baselines('doc-grid-none.docx');

        expect(Math.abs(firstGridded! - 800.89)).toBeLessThan(TOLERANCE_PT);
        expect(Math.abs((firstPlain! - firstGridded!) - 3.25)).toBeLessThan(TOLERANCE_PT);
    });

    const stepsOf = async (name: string): Promise<number[]> => {
        const marks = await baselines(name);

        return marks.slice(1).map(
            (value, index) => Math.round((marks[index]! - value) * 100) / 100);
    };

    it('ignores a `default` grid, whatever pitch it states', async () => {
        // Measured: `w:type="default"` with a 360-twip pitch stepped 11.5,
        // the font's own line. Only `lines` and `linesAndChars` typeset.
        expect(await stepsOf('doc-grid-default.docx')).toEqual([11.5, 11.5, 11.5]);
    });

    it('treats `linesAndChars` as `lines` for Latin text', async () => {
        // Same 18pt step, and its `w:charSpace` of 400 changed nothing —
        // character gridding is for the scripts this engine does not shape.
        expect(await stepsOf('doc-grid-lines-and-chars.docx')).toEqual([18, 18, 18]);
    });

    it('lets a paragraph’s stated spacing WIN over the grid', async () => {
        // An exact 280-twip line under an 18pt grid stepped 14, not 18: the
        // grid is what a paragraph falls back to, not what it is held to.
        expect(await stepsOf('doc-grid-with-spacing.docx')).toEqual([14, 14, 14]);
    });

    it('lets a paragraph’s own spacing beat the grid', async () => {
        // The grid is a fallback, not an override: a paragraph stating
        // `w:spacing` keeps what it stated. Asserted on the MODEL, because no
        // fixture here states both.
        const opened = await openWordFile(file('doc-grid-lines.docx'), FONTS);
        const [block] = opened.document.sections[0]!.blocks;

        expect(isTable(block!) ? undefined : block!.style.lineRule).toBe('grid');
    });
});

describe('w:textDirection, end to end', () => {
    // Measured through LibreOffice. `cell-text-turned.docx` is one turned cell
    // beside three upright paragraphs; `-wrapped` is a longer turned run beside
    // four. Both printed every glyph with the text matrix [0 1 -1 0] — a
    // quarter turn anticlockwise — which is what `btLr` means.
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;
    const upFromFoot = (px: number): number => PAGE_HEIGHT_PT - pt(px);

    const pageOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections)[0]!;
    };
    const turnedCellOf = async (name: string) => (await pageOf(name)).rows[0]!.cells[0]!;

    it('turns the cell anticlockwise, as btLr asks', async () => {
        expect((await turnedCellOf('cell-text-turned.docx')).turn).toBe('ccw');
    });

    it('gives a turned line the ROW’s height to break at, not the column’s width', async () => {
        // Six characters would not fit upright in a 2000-twip column, and the
        // printed page put all six on one turned line.
        const cell = await turnedCellOf('cell-text-turned.docx');

        expect(cell.lines).toHaveLength(1);
        expect(cell.lines[0]!.line.pieces.map((piece) => piece.text).join('')).toBe('abcdef');
    });

    it('breaks a longer turned run where LibreOffice broke it', async () => {
        // The proof that our row height IS LibreOffice's: break the same run at
        // a height that differs by more than a word and the words move lines.
        // Printed: `one two` / `three four` / `five six`.
        const cell = await turnedCellOf('cell-text-turned-wrapped.docx');

        expect(cell.lines.map((placed) =>
            placed.line.pieces.map((piece) => piece.text).join('').trimEnd()))
            .toEqual(['one two', 'three four', 'five six']);
    });

    it('stacks turned lines across the cell, one printed x each', async () => {
        // Printed x: 43.15, 55.35, 67.55 — a 12.2pt line pitch running left to
        // right, because the anticlockwise turn puts later lines to the RIGHT.
        const rendered = renderPage(await pageOf('cell-text-turned-wrapped.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => ['one', 'three', 'five'].includes(op.text.trim()));

        expect(rendered.map((op) => op.turn)).toEqual(['ccw', 'ccw', 'ccw']);
        expect(rendered.map((op) => Math.round(pt(op.xPx) * 10) / 10)).toEqual([43.1, 55.3, 67.5]);
    });

    it('starts every turned line at the row’s foot, climbing the page', async () => {
        // Printed: all three lines begin at y=751.64, measured up from the foot
        // of the page — a turned line runs bottom to top, so its start is the
        // BOTTOM of the cell and each glyph after it is higher.
        const rendered = renderPage(await pageOf('cell-text-turned-wrapped.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => ['one', 'three', 'five'].includes(op.text.trim()));

        for (const op of rendered) {
            expect(Math.abs(upFromFoot(op.yPx) - 751.64)).toBeLessThan(0.5);
        }
    });

    it('draws the single-line cell where LibreOffice drew it', async () => {
        // Printed: x=43.00, first glyph at y=766.64 up from the page foot.
        const [drawn] = renderPage(await pageOf('cell-text-turned.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'abcdef' === op.text);

        expect(Math.abs(pt(drawn!.xPx) - 43.00)).toBeLessThan(0.5);
        expect(Math.abs(upFromFoot(drawn!.yPx) - 766.64)).toBeLessThan(0.5);
    });

    it('does not let a turned cell grow its row', async () => {
        // The turned run is far longer than the row is tall; the row is still
        // exactly as tall as the four upright paragraphs beside it asked.
        const page = await pageOf('cell-text-turned-wrapped.docx');
        const [turned, upright] = page.rows[0]!.cells;

        expect(turned!.heightPx).toBe(upright!.heightPx);
        expect(upright!.lines).toHaveLength(4);
    });

    it('leaves a cell upright when no upright neighbour sets a height', async () => {
        // There is no measurement to copy here: LibreOffice prints
        // `cell-text-turned-alone.docx` as a blank page — no text, not even the
        // table's rules. Upright is this engine's answer to a
        // question the reference declines to answer, and it keeps the text.
        const cell = await turnedCellOf('cell-text-turned-alone.docx');

        expect(cell.turn).toBeUndefined();
        expect(cell.lines.length).toBeGreaterThan(0);
    });

    it('advances along a turned line UP the page, not down', async () => {
        // Printed line one is `one two`: `one` begins at y=751.64 and `two` at
        // y=769.39, further UP. Advancing the other way would keep the line's
        // start right and run the rest of it off the bottom of the cell.
        const words = renderPage(await pageOf('cell-text-turned-wrapped.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => ['one', 'two'].includes(op.text.trim()));

        expect(words).toHaveLength(2);
        expect(upFromFoot(words[1]!.yPx)).toBeGreaterThan(upFromFoot(words[0]!.yPx));
        expect(Math.abs(upFromFoot(words[1]!.yPx) - 769.39)).toBeLessThan(0.5);
        expect(words[0]!.xPx).toBe(words[1]!.xPx);
    });

    it('carries the turn into the SVG, as a quarter turn about the anchor', async () => {
        // The whole point of the turn reaches the page through this one seam.
        const opened = await openWordFile(file('cell-text-turned.docx'), FONTS);
        const svg = renderPageToSvg(layoutSections(opened.document.sections)[0]!);

        expect(svg).toMatch(/rotate\(-90 /);
        expect(svg).not.toMatch(/rotate\(90 /);
    });

    /**
     * Where the turned run of one of the `-inset`/`-indent` fixtures starts.
     *
     * The word is `abcdef` and the first LINE of it is what starts at the foot:
     * both fixtures narrow the line enough that LibreOffice chops the word, and
     * this engine chops it in the same place. Matching on the whole word found
     * nothing once that landed — and had been hiding the chop in the printed
     * page all along.
     */
    const startOfTurnedRun = async (name: string): Promise<number> => {
        const [drawn] = renderPage(await pageOf(name)).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => op.text.trim().startsWith('abc'));

        return upFromFoot(drawn!.yPx);
    };

    it('starts the turned line a BOTTOM CELL MARGIN above the cell’s foot', async () => {
        // `-inset` sets w:tcMar top and bottom to 283 twips (14.15pt). Printed:
        // the row's bottom rule at y=734.889 and the run starting at 749.69 —
        // 14.8 above it, the margin plus the same 0.65 the unmargined fixture
        // shows between a rule and the text it bounds.
        expect(Math.abs(await startOfTurnedRun('cell-text-turned-inset.docx') - 749.69))
            .toBeLessThan(0.5);
    });

    it('starts it an INDENT above the foot, because the indent runs along the height', async () => {
        // `-indent` sets w:ind left to 200 twips (10pt) on the turned
        // paragraph. Printed: bottom rule at 763.189, run starting at 773.84 —
        // 10.65 above it. An indent in a turned cell moves text along the row's
        // height, not across its width.
        expect(Math.abs(await startOfTurnedRun('cell-text-turned-indent.docx') - 773.84))
            .toBeLessThan(0.5);
        // And the chop itself, which the printed page shows and this test used
        // to look straight past: `abcde` on the first turned line and `f` on
        // the second, whose baseline is 12.20 further across at x=55.35.
        const [second] = renderPage(await pageOf('cell-text-turned-indent.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'f' === op.text.trim());

        expect(Math.abs(pt(second!.xPx) - 55.35)).toBeLessThan(0.5);
    });

    it('will not grow the row for turned lines that overflow the cell', async () => {
        // `-overflow` is 18 words turned in a 900-twip cell beside three
        // upright paragraphs. Printed: the row stayed 37.6pt — exactly the
        // three paragraphs — and the FOURTH turned line was drawn at x=78.70,
        // past the cell's right edge at 72.90, with the rest of the text drawn
        // nowhere at all. LibreOffice overflows and then gives up; it never
        // makes the row taller to win back the room.
        const page = await pageOf('cell-text-turned-overflow.docx');
        const [turned, upright] = page.rows[0]!.cells;

        expect(turned!.heightPx).toBe(upright!.heightPx);
        expect(Math.abs(pt(page.rows[0]!.heightPx) - 36.6)).toBeLessThan(1.5);
        // Where we differ: the fourth line is the one LibreOffice drew outside
        // the cell, and we stop at the cell's edge instead. Both lose the same
        // fourteen words after it; neither is a document anyone would ship.
        expect(turned!.lines.map((placed) =>
            placed.line.pieces.map((piece) => piece.text).join('').trimEnd()))
            .toEqual(['alpha', 'bravo', 'charlie']);
    });

    it('will not grow the row even when the turned cell is WIDE and its neighbour short', async () => {
        // The case that separates `the row's height' from `whatever the turned
        // text needs': a 4000-twip turned cell of eighteen words beside a
        // single upright paragraph. Printed: the row is 13.2pt — that one
        // paragraph — and the turned text ran to SEVENTEEN lines, the last at
        // x=233.70, past the cell's own right edge at 227.90. A turned cell
        // spills sideways out of the table before it makes its row any taller.
        const page = await pageOf('cell-text-turned-wide.docx');
        const [turned, upright] = page.rows[0]!.cells;

        // How many lines each engine makes of it is not worth asserting: at a
        // 12.2pt line length LibreOffice put about two glyphs on a line and
        // dropped the rest of each word, and we take the documented overflow
        // path instead. Neither is output anyone would ship. The row height is
        // the measurement, and both engines agree on it.
        expect(turned!.heightPx).toBe(upright!.heightPx);
        expect(Math.abs(pt(page.rows[0]!.heightPx) - 13.2)).toBeLessThan(1.5);
    });

    it('paints a turned highlight as the upright box with its sides swapped', async () => {
        // Printed: `re 33.600 763.739 12.150 14.250` — 12.15 across, which is
        // the LINE's height, by 14.25 along, which is the highlighted run's own
        // length. A quarter turn leaves a rectangle axis-aligned.
        const [box] = renderPage(await pageOf('cell-text-turned-decorated.docx')).ops
            .filter((op): op is RectOp => 'rect' === op.kind)
            .filter((op) => '#ffff00' === op.fill.toLowerCase());

        expect(Math.abs(pt(box!.xPx) - 33.600)).toBeLessThan(0.5);
        expect(Math.abs(upFromFoot(box!.yPx + box!.heightPx) - 763.739)).toBeLessThan(0.5);
        expect(Math.abs(pt(box!.widthPx) - 12.150)).toBeLessThan(0.5);
        expect(Math.abs(pt(box!.heightPx) - 14.250)).toBeLessThan(0.5);
    });

    it('rules a turned underline BESIDE the baseline, running with the text', async () => {
        // LibreOffice draws this one inside the rotation: `0 -1.1 l 14.2 -1.1`,
        // the run's own length at the font's offset from the baseline. Turned,
        // that offset lands on X — the descender side, which for an
        // anticlockwise turn is to the RIGHT of the text.
        const drawn = renderPage(await pageOf('cell-text-turned-decorated.docx')).ops
            .filter((op): op is LineOp => 'line' === op.kind)
            .filter((op) => op.x1Px === op.x2Px && Math.abs(pt(op.x1Px) - 44.25) < 1);

        expect(drawn).toHaveLength(1);
        expect(Math.abs(pt(Math.abs(drawn[0]!.y2Px - drawn[0]!.y1Px)) - 14.2)).toBeLessThan(0.5);
        // It climbs with the text rather than running back down it.
        expect(drawn[0]!.y2Px).toBeLessThan(drawn[0]!.y1Px);
    });

    it('turns tbRl the OTHER way, as a mirror of btLr', async () => {
        // Printed with the matrix [-0 -1 1 -0] — a quarter turn clockwise,
        // where btLr printed [0 1 -1 0]. The two are mirror images, not the
        // same placement with a different sign.
        const cell = await turnedCellOf('cell-text-turned-clockwise.docx');

        expect(cell.turn).toBe('cw');
    });

    it('starts a clockwise line at the cell’s HEAD and runs it down', async () => {
        // Printed: the run begins at y=800.24, just under the row's top rule at
        // 800.789, and descends to 775.84. Anticlockwise starts at the foot and
        // climbs; reusing that placement put clockwise text below the cell.
        const [drawn] = renderPage(await pageOf('cell-text-turned-clockwise.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'abc' === op.text.trim());

        expect(drawn!.turn).toBe('cw');
        expect(Math.abs(upFromFoot(drawn!.yPx) - 800.24)).toBeLessThan(0.5);
    });

    it('stacks clockwise lines from the cell’s RIGHT edge', async () => {
        // Printed: the baseline at x=112.50, in a cell whose inner right edge
        // is 122.05 — the line hangs off the right, where btLr hangs off the
        // left. Measuring from the left edge would have put it at 43.15.
        const [drawn] = renderPage(await pageOf('cell-text-turned-clockwise.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'abc' === op.text.trim());

        expect(Math.abs(pt(drawn!.xPx) - 112.50)).toBeLessThan(0.5);
    });

    it('turns the clockwise cell the other way in the SVG too', async () => {
        const opened = await openWordFile(file('cell-text-turned-clockwise.docx'), FONTS);
        const svg = renderPageToSvg(layoutSections(opened.document.sections)[0]!);

        expect(svg).toMatch(/rotate\(90 /);
        expect(svg).not.toMatch(/rotate\(-90 /);
    });

    it('advances a clockwise line DOWN the page, mirroring the other turn', async () => {
        // Printed: `abc` at y=800.24 and `def` at 783.74, descending. An
        // anticlockwise line climbs; sharing one advance between the two turns
        // sends one of them out of its cell.
        const words = renderPage(await pageOf('cell-text-turned-clockwise.docx')).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => ['abc', 'def'].includes(op.text.trim()));

        expect(words).toHaveLength(2);
        expect(upFromFoot(words[1]!.yPx)).toBeLessThan(upFromFoot(words[0]!.yPx));
        expect(Math.abs(upFromFoot(words[1]!.yPx) - 783.74)).toBeLessThan(0.5);
        expect(words[0]!.xPx).toBe(words[1]!.xPx);
    });

    it('draws a boxed paragraph’s BOX in a turned cell, turned with it', async () => {
        // A quarter turn leaves a rectangle axis-aligned, so the box is the
        // upright one with its axes exchanged. Printed: the rules of
        // a box round a single turned line at x=77.80 and 98.25, from y=717.94
        // to 762.39 — 20.45 across, which is the line's 11.50 with `w:space`
        // 4pt and half a point of rule either side, by 44.45 along, which is
        // the row's 35.50 with the same at each end.
        //
        // `placeRow` handed a turned cell's `paragraphBorders` back EMPTY until
        // this row, so the box was placed and never drawn — the defect the
        // placed-and-never-drawn guard exists for, in the one place that guard
        // cannot see, because a cell's boxes are not the page's.
        const page = await pageOf('cell-turned-extras.docx');
        const [box] = page.rows[0]!.cells[0]!.paragraphBorders;

        expect(box).toBeDefined();
        expect(Math.abs(pt(box!.leftPx) - 77.80)).toBeLessThan(0.5);
        expect(Math.abs(pt(box!.rightPx) - 98.25)).toBeLessThan(0.5);
        // The along-axis edges carry the row's own height, and ours is 34.50
        // where LibreOffice prints 35.50 — the table rule-room difference this
        // arc has measured before, not a fault in the turn.
        expect(Math.abs(upFromFoot(box!.bottomPx) - 717.94)).toBeLessThan(1);
        expect(Math.abs(upFromFoot(box!.topPx) - 762.39)).toBeLessThan(1);
        // And it reaches the ink: the box's left rule is at an x no table rule
        // of this fixture shares.
        expect(renderPage(page).ops.filter((op): op is LineOp => 'line' === op.kind)
            .some((op) => Math.abs(pt(op.x1Px) - 77.75) < 0.1 && op.x1Px === op.x2Px)).toBe(true);
    });

    it('charges a turned line a SHAPE’s height, as it does a picture’s', async () => {
        // Measured: a 36x18 shape in a turned cell started the text
        // after it 18.05 above the row's foot — its height, not the 36.00 of
        // its width, which is the rule measured for pictures and which
        // `turnPictures` was applying to pictures alone.
        //
        // ACROSS the line it is still wrong and now stated: LibreOffice printed
        // the text at x=108.95, 31.70 from the cell's inner edge, where we put
        // it at 113.25 — the shape's whole 36.00, standing on the baseline.
        // 31.70 is neither that nor the 22.69 a picture's centring would give,
        // and one measurement is not a rule.
        const page = await pageOf('cell-turned-extras.docx');
        const [text] = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'sh' === op.text.trim());

        expect(Math.abs(upFromFoot(text!.yPx) - 645.04)).toBeLessThan(0.5);
        expect(Math.round(pt(text!.xPx) * 100) / 100).toBe(113.25);
    });

    it('charges nothing for a shape too LONG for its turned line', async () => {
        // The fourth table's shape is 36x40 in a row 35.5pt tall. Printed: the
        // text after it begins 0.60 above the row's foot, which is where a
        // turned line with nothing reserved begins — the same "gives up rather
        // than reserving what it has" measured for pictures,
        // now measured for shapes rather than carried over on the strength of
        // the resemblance. A mutation charging what is LEFT survived until this
        // table existed.
        //
        // Across the line it is 63.45 from the cell's inner edge where we put
        // it at 113.25 — 27.45 out, and a second unexplained cross-axis number
        // for shapes beside the 4.30 above. Two measurements, no rule.
        // The 0.60 is LibreOffice's own inset — every turned line gets it, and
        // it is not room the shape asked for (the picture rule reads the same
        // fixture family the same way). What is asserted is the RESERVATION:
        // nought.
        const page = await pageOf('cell-turned-extras.docx');
        const row = page.rows[3]!;
        const [text] = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'big' === op.text.trim());
        const reservedPx = upFromFoot(text!.yPx) - upFromFoot(row.yPx + row.heightPx);

        expect(Math.abs(reservedPx)).toBeLessThan(0.5);
    });

    it('draws NOTHING for a table nested in a turned cell, where LibreOffice draws it', async () => {
        // Measured and NOT built: LibreOffice prints a nested table
        // inside a turned cell — its `N1` turned at x=87.65, inside rules
        // 13.50 across by 36.50 along. We drop the whole thing.
        //
        // The difference from the paragraph box above is a tree: turning a row
        // means turning every cell, line, rule and nested row it holds, where a
        // box is one rectangle with its axes exchanged. Pinned so the day it is
        // built this test fails.
        const cell = (await pageOf('cell-turned-extras.docx')).rows[1]!.cells[0]!;

        expect(cell.turn).toBe('ccw');
        expect(cell.rows).toEqual([]);
        expect(cell.lines).toEqual([]);
    });

    it('draws NO picture in a turned cell, which is what LibreOffice does', async () => {
        // Not a gap — parity. The same fixture with the turn taken off prints
        // the picture at 18.00 x 9.00, so the drawing is valid and the media
        // part intact; with `btLr` on the cell, LibreOffice prints the text and
        // no image at all. Verified with tools/probes/pdf-images.py, which
        // found the control and found nothing here.
        const page = await pageOf('cell-picture-turned.docx');
        const drawn = renderPage(page).ops.filter((op) => 'image' === op.kind);

        expect(drawn).toHaveLength(0);
        // The picture is still a piece of the line — it is undrawn, not gone.
        expect(page.rows[0]!.cells[0]!.lines[0]!.line.pieces
            .some((piece) => undefined !== piece.image)).toBe(true);
    });

    it('charges the turned line the picture’s HEIGHT, as LibreOffice does', async () => {
        // Measured over five sizes in one conversion: ALONG a turned
        // line a picture reserves its height. 18x9 and 36x9 both reserved
        // 9.00; 18x18 and 9x18 both 18.00 — doubling the width moved nothing.
        //
        // The constant comes from the same document: 36x36 in a 35.5pt row
        // cannot fit, reserves nothing, and still started its text 0.60 above
        // the row's foot. So this 18x9 picture's printed 9.60 is that 0.60
        // plus a 9.00 reservation, and 9.00 is the number to match.
        //
        // ACROSS the line it is CENTRED rather than reserved, which is a
        // different rule in a different place — see the cross-axis tests below.
        const page = await pageOf('cell-picture-turned.docx');
        const foot = upFromFoot(page.rows[0]!.yPx + page.rows[0]!.heightPx);
        const [text] = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'ab' === op.text.trim());

        expect(Math.abs(foot - 722.389)).toBeLessThan(0.5);
        expect(Math.abs(upFromFoot(text!.yPx) - foot - 9.00)).toBeLessThan(0.5);
    });

    it('leaves an UPRIGHT cell’s picture alone, advancing by its width', async () => {
        // The control that proved that null, kept as a fixture because it
        // is the only picture in a table CELL that is not turned. Printed: the
        // picture at 18.00 x 9.00 with its left edge on the cell's inner edge
        // at 77.25, and the text after it at 95.35 — an advance of 18.10, the
        // picture's WIDTH. Exchanging a picture's sides here would move it.
        const page = await pageOf('cell-picture-upright.docx');
        const [text] = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => 'ab' === op.text.trim());

        expect(Math.abs(pt(text!.xPx) - 95.35)).toBeLessThan(0.5);
        // And unlike a turned cell, this one really is drawn.
        expect(renderPage(page).ops.filter((op) => 'image' === op.kind)).toHaveLength(1);
    });

    it('reserves each picture’s height across five sizes, and nothing when it cannot fit', async () => {
        // The turned-picture fixture: five turned cells, five picture sizes,
        // one conversion. Printed reservations along the line, each measured as the
        // text's start above its row's foot less the 0.60 an unreserved turned
        // line gets: 18x9 -> 9.00, 36x9 -> 9.00, 18x18 -> 18.00, 9x18 -> 18.00.
        // Width doubles without moving any of them.
        //
        // The fifth is 36x36 in a 35.5pt row. It cannot fit, and LibreOffice
        // stops charging for it rather than overflowing: its text starts 0.60
        // above the foot, exactly like a line with no picture at all.
        const page = await pageOf('cell-picture-turned-sizes.docx');
        const reserved = page.rows.map((row) => {
            const foot = upFromFoot(row.yPx + row.heightPx);
            const [text] = renderPage(page).ops
                .filter((op): op is TextOp => 'text' === op.kind)
                .filter((op) => 'ab' === op.text.trim()
                    && upFromFoot(op.yPx) > foot - 1
                    && upFromFoot(op.yPx) < foot + row.heightPx * 72 / 96 + 1);

            return undefined === text ? null : Math.round((upFromFoot(text.yPx) - foot) * 100) / 100;
        });

        // Four sizes, exact, with width varied to prove it is the height.
        expect(reserved.slice(0, 4)).toEqual([9, 9, 18, 18]);
        // The fifth cannot fit: 36x36 in a 35.5pt row. LibreOffice gives up
        // rather than reserving what it has — its text starts 0.60 above the
        // foot, exactly like a line with no picture — so nothing is charged.
        expect(reserved[4]).toBe(0);
    });

    it('centres a turned picture ACROSS the line rather than standing it on the baseline', async () => {
        // The other axis of the same five pictures, from the same printed
        // page. The cell's inner edge is at 77.25, so the printed
        // baselines 90.95, 99.95, 90.95, 86.65 and 122.65 are 13.70, 22.70,
        // 13.70, 9.40 and 45.40 across, for pictures 18, 36, 18, 9 and 36 wide.
        //
        // Half the width plus a constant, and the constant is half the ascent:
        // (9.40 + 18) / 2 is 13.70 and (9.40 + 36) / 2 is 22.70, against the
        // 9.40 of the ascent alone where the picture is narrower than it and
        // asks for no room. Standing the picture on the baseline, which is
        // what an upright line does and what we did until now, puts the first
        // two at 18.00 and 36.00 — out by more than four points.
        //
        // The fifth is the 36x36 that cannot fit its line, and it is the one
        // case that is NOT centred: 45.40 is the whole picture and then the
        // text's own ascent, the text hanging below it.
        const page = await pageOf('cell-picture-turned-sizes.docx');
        const drawn = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);
        const across = page.rows.map((row) => {
            const [text] = drawn.filter((op) => 'ab' === op.text.trim()
                && op.yPx >= row.yPx - 1 && op.yPx <= row.yPx + row.heightPx + 1);

            return pt(text!.xPx);
        });

        expect(across).toHaveLength(5);
        for (const [index, printed] of [90.95, 99.95, 90.95, 86.65, 122.65].entries()) {
            expect(Math.abs(across[index]! - printed)).toBeLessThan(0.5);
        }
    });

    it('keeps the descender below the PICTURE’s foot, not below the baseline', async () => {
        // A single-line cell cannot see a line's BOX at all — it shows where
        // one baseline landed and no more, which is why an earlier probe read
        // three baselines and could not tell a placement from a box. Each
        // turned cell here wraps over two lines, so the second baseline
        // reports the first line's whole height.
        //
        // Printed in one conversion, four turned cells whose inner edge is at
        // 77.25 — no picture, then 18x9, 36x9 and 36x36:
        //
        //   no picture     86.65, 98.15
        //   18 across      90.95, 106.75
        //   36 across      99.95, 124.75
        //   36, too long  122.65, 134.15
        //
        // The second line of each carries no picture, so it is a plain line
        // whose baseline sits one ascent — 9.40, which the first row prints
        // outright — below the first line's foot. That makes the first line's
        // box 11.50, 20.10, 38.10 and 47.50: the PICTURE plus one descender,
        // even though the baseline sits at 13.70 and 22.70 inside it.
        // LibreOffice keeps the descender under the picture's own foot rather
        // than under the baseline, so a centred picture overhangs its line.
        // Sizing the box from the baseline instead — the obvious reading of
        // "centred", and the one this fixture exists to refuse — makes it
        // 15.80 and 24.80, wrong by a third.
        const page = await pageOf('cell-picture-turned-wrapped.docx');

        // Count what came back: five tables in, five rows out. The
        // fifth is an upright cell, added later, and is read there.
        expect(page.rows).toHaveLength(5);

        page.rows.slice(0, 4).forEach((row, index) => {
            const first = row.cells[0]!.lines[0]!;

            expect(Math.abs(pt(first.heightPx) - [11.50, 20.10, 38.10, 47.50][index]!))
                .toBeLessThan(0.5);
            expect(Math.abs(pt(first.xPx + first.baselinePx) - [86.65, 90.95, 99.95, 122.65][index]!))
                .toBeLessThan(0.5);
        });
    });

    it('gives the picture’s box to ITS line and leaves the next one plain', async () => {
        // The second line of each cell carries no picture, so it is an
        // ordinary 11.50 line: printed at 98.15, 106.75, 124.75 and 134.15.
        //
        // This test was written as a DIVERGENCE — ours were 98.13, 111.06,
        // 138.06 and 170.13, because a line's box was measured per paragraph
        // and the picture on the first line made every line after it as tall.
        // Making the box per line flipped it to parity without the test being
        // touched, which is the whole point of pinning our own numbers rather
        // than shrugging at the gap.
        const page = await pageOf('cell-picture-turned-wrapped.docx');
        const second = page.rows.slice(0, 4).map((row) => pt(row.cells[0]!.lines[1]!.xPx
            + row.cells[0]!.lines[1]!.baselinePx));

        for (const [index, printed] of [98.15, 106.75, 124.75, 134.15].entries()) {
            expect(Math.abs(second[index]! - printed)).toBeLessThan(0.5);
        }
    });
});

describe('a line box belongs to its LINE, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;
    const upFromFoot = (px: number): number => PAGE_HEIGHT_PT - pt(px);

    /**
     * Every baseline drawn on the page, top to bottom, as the PDF reports one.
     *
     * A line is drawn piece by piece and LibreOffice shows it in one go, so the
     * pieces are gathered back up by the baseline they share.
     */
    const baselinesOf = async (name: string): Promise<
        { text: string; xPx: number; yPx: number }[]
    > => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const byBaseline = new Map<number, { text: string; xPx: number }>();

        for (const op of renderPage(page).ops) {
            if ('text' !== op.kind || '' === op.text.trim()) {
                continue;
            }
            const so_far = byBaseline.get(op.yPx);

            byBaseline.set(op.yPx, {
                text: (so_far?.text ?? '') + op.text,
                xPx: Math.min(so_far?.xPx ?? op.xPx, op.xPx),
            });
        }

        return [...byBaseline.entries()]
            .map(([yPx, line]) => ({ ...line, text: line.text.trim(), yPx: upFromFoot(yPx) }));
    };

    it('makes the picture’s own line taller and every other line plain', async () => {
        // `inline-picture-wrapped.docx` is three narrow paragraphs, printed in
        // one conversion: an 18x36 picture on the FIRST line of one,
        // the same picture in the MIDDLE of the next, and neither in the third.
        //
        // Printed baselines, measured up from the foot of the page:
        //
        //   gap0    760.49
        //   line 1  722.39   38.10 below gap0 — the picture's line
        //   line 2  710.89   11.50 below it — an ordinary line
        //   gap1    699.39
        //   line 1  687.89   11.50 — ordinary, the picture is not on it
        //   line 2  649.79   38.10 — the tall line MOVED with the picture
        //   gap2    638.29   11.50
        //   line 1  626.79 · line 2  615.29 · end  603.79
        //
        // So the picture makes ITS line taller and no other. Measuring the box
        // per paragraph — what this engine used to do — spaces every
        // line of a paragraph by its largest thing, which puts the second line
        // of the first paragraph 27pt low and pushes the rest of the page down
        // with it.
        const printed = [
            ['gap0', 760.49], ['aa bb cc dd ee ff', 722.39], ['gg hh ii jj kk ll mm nn', 710.89],
            ['gap1', 699.39], ['pp qq rr ss tt uu vv ww', 687.89], ['xx yy zz a1 b1 c1', 649.79],
            ['gap2', 638.29], ['aa bb cc dd ee ff gg', 626.79], ['hh ii jj kk ll mm nn', 615.29],
            ['end', 603.79],
        ] as const;
        const ours = await baselinesOf('inline-picture-wrapped.docx');

        // Count what came back before reading it.
        expect(ours).toHaveLength(printed.length);
        printed.forEach(([text, y], index) => {
            expect(ours[index]!.text).toBe(text);
            expect(Math.abs(ours[index]!.yPx - y)).toBeLessThan(0.5);
        });
    });

    it('gives a table CELL’s line its own box too, wherever the picture lands', async () => {
        // The fifth table of `cell-picture-turned-wrapped.docx` is an UPRIGHT
        // cell whose 18x36 picture is pushed onto the SECOND line by the text
        // in front of it. That is the one case a cell fixture with its picture
        // on the first line cannot see — there the paragraph's nominal height
        // and the first line's own are the same number — and it is why a
        // mutation stacking a cell's lines by the nominal height survived the
        // first pass. The table was added to kill it.
        //
        // Printed: the first line plain at y=555.99, the second at 517.89. The
        // step is 38.10, the picture's box, and it belongs to the line the
        // picture landed on rather than to the paragraph's first.
        const opened = await openWordFile(file('cell-picture-turned-wrapped.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const lines = page.rows[4]!.cells[0]!.lines;

        expect(lines).toHaveLength(2);
        expect(Math.abs(upFromFoot(lines[0]!.yPx + lines[0]!.baselinePx) - 555.99)).toBeLessThan(0.5);
        expect(Math.abs(upFromFoot(lines[1]!.yPx + lines[1]!.baselinePx) - 517.89)).toBeLessThan(0.5);
        // And the picture stands on that second baseline, at the cell's inner
        // edge, with the text after it one picture's width along at 95.35.
        const [drawn] = renderPage(page).ops.filter((op): op is ImageOp => 'image' === op.kind);

        expect(Math.abs(upFromFoot(drawn!.yPx + drawn!.heightPx) - 517.94)).toBeLessThan(0.5);
        expect(Math.abs(pt(drawn!.xPx) - 77.25)).toBeLessThan(0.5);
    });

    it('walks a float’s band by the lines’ real heights, not by counting them', async () => {
        // `float-tall-first-line.docx` puts the same 72x72 float and the same
        // forty words in two paragraphs of the same 201pt measure. The second
        // begins with an inline 18x36 picture, so its first line is 38.10 tall
        // where every other line in the document is 11.50.
        //
        // Printed — a displaced line starts at x=153.10, a clear one
        // at 72.10:
        //
        //   float only     153.10 x7, then 72.10  — seven lines beside it
        //   tall first     171.10, 153.10 x3, then 72.10 x3
        //
        // The first line of the second paragraph starts at 171.10 because the
        // picture takes the first 18pt of it. FOUR lines are displaced there
        // against seven here, and the difference is entirely the tall line
        // eating 38.10 of a 72pt band.
        //
        // Walking the band by line INDEX — `startY + N * lineHeight` with the
        // paragraph's nominal height, which is what this engine did — steps
        // 38.10 a line through that second paragraph and clears the float
        // after two, so lines three and four are broken at the full 201pt and
        // drawn over the picture.
        const lines = await baselinesOf('float-tall-first-line.docx');
        const printed = [
            72.10,                                              // gap0
            153.10, 153.10, 153.10, 153.10, 153.10, 153.10, 153.10, 72.10,
            72.10,                                              // gap1
            171.10, 153.10, 153.10, 153.10, 72.10, 72.10, 72.10,
            72.10,                                              // end
        ];

        // Count what came back: eighteen lines printed, eighteen laid out.
        expect(lines).toHaveLength(printed.length);
        printed.forEach((x, index) => {
            expect(Math.abs(pt(lines[index]!.xPx) - x)).toBeLessThan(0.5);
        });
        // The words a displaced line holds are the proof it was broken at the
        // narrow width rather than merely drawn there.
        expect(lines[10]!.text).toBe('w01 w02 w03 w04');
        expect(lines[14]!.text).toBe('w20 w21 w22 w23 w24 w25 w26 w27 w28');
    });

    it('stands the picture on the baseline of the line it landed on', async () => {
        // Printed by pdf-images.py: the two pictures at y=722.44 and 649.84,
        // which are their own lines' baselines — an inline picture stands on
        // the baseline whichever line it ends up on.
        const opened = await openWordFile(file('inline-picture-wrapped.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const feet = renderPage(page).ops
            .filter((op): op is ImageOp => 'image' === op.kind)
            .map((op) => upFromFoot(op.yPx + op.heightPx));

        expect(feet).toHaveLength(2);
        expect(Math.abs(feet[0]! - 722.44)).toBeLessThan(0.5);
        expect(Math.abs(feet[1]! - 649.84)).toBeLessThan(0.5);
    });
});

describe('w:spacing against what the line actually holds, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    /** Every body baseline, top to bottom, measured up from the foot. */
    const linesOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return page.lines.map((line) => ({
            text: line.line.pieces.map((piece) => piece.text).join('').trim(),
            heightPx: pt(line.heightPx),
            baselineY: PAGE_HEIGHT_PT - pt(line.yPx + line.baselinePx),
        }));
    };

    it('lets a picture grow an atLeast or a proportional line, and clips it under exact', async () => {
        // `line-spacing-picture.docx`: the same 18x36 picture on the first line
        // of four paragraphs that differ only in `w:spacing`. Printed,
        // thirteen baselines up from the foot:
        //
        //   none      760.49 | 722.39  710.89
        //   exact 12  699.39 | 687.69  675.69
        //   atLeast12 663.89 | 625.79  613.79
        //   auto 1.5  602.29 | 564.19  546.94   then end at 529.69
        //
        // `exact` keeps its 12.00 and lets the picture overflow into the
        // paragraph above — 11.70 from the line before it, where an unspaced
        // picture line takes 38.10. Nothing else clips: `atLeast` grows to the
        // picture's own 38.10, and `auto` to 43.85, which is that box plus the
        // 5.75 of leading its multiple adds to every line of the paragraph.
        //
        // Before this was measured we gave the `auto` line a 17.25 box with its
        // baseline 36.00 inside it — the baseline BELOW the box's own foot, so
        // the second line was drawn above the first.
        const printed = [
            760.49, 722.39, 710.89,
            699.39, 687.69, 675.69,
            663.89, 625.79, 613.79,
            602.29, 564.19, 546.94, 529.69,
        ];
        const ours = await linesOf('line-spacing-picture.docx');

        expect(ours).toHaveLength(printed.length);
        printed.forEach((y, index) => {
            expect(Math.abs(ours[index]!.baselineY - y)).toBeLessThan(0.5);
        });
        // The four first lines, stated as boxes rather than as positions.
        expect([1, 4, 7, 10].map((index) => Math.round(ours[index]!.heightPx * 10) / 10))
            .toEqual([38.1, 12, 38.1, 43.9]);
    });

    it('adds a multiple’s leading to each line’s own box, not to the paragraph’s', async () => {
        // `line-spacing-multiple-mixed.docx` at 1.5 lines over a paragraph
        // whose tallest run is 23.00, so the leading is 11.50. Printed: the
        // 10pt first line 23.00 deep and the line carrying the 20pt word 34.50
        // — each its own box plus that same 11.50, and neither one a multiple
        // of anything. Baselines 687.89 and 655.54, with `end` at 630.39.
        //
        // Taking the multiple against the LINE would give 17.25 and 34.50, and
        // against the PARAGRAPH 34.50 twice. The mixed sizes are what tell the
        // three apart; a uniform paragraph cannot.
        const ours = await linesOf('line-spacing-multiple-mixed.docx');
        const printed = [760.49, 722.39, 710.89, 699.39, 687.89, 655.54, 630.39];

        expect(ours).toHaveLength(printed.length);
        printed.forEach((y, index) => {
            expect(Math.abs(ours[index]!.baselineY - y)).toBeLessThan(0.5);
        });
        expect([4, 5].map((index) => Math.round(ours[index]!.heightPx * 10) / 10))
            .toEqual([23, 34.5]);
    });

    it('lets a paragraph come OFF the grid with w:snapToGrid', async () => {
        // Word writes the off form freely, so a gridded document is likely to
        // hold some, and a reader that ignores it lays them on a grid the file
        // says to keep them off. Printed under an 18pt grid: the
        // paragraph that says no steps 11.50 — the font's own line — where its
        // neighbour on the grid steps 18.00, and its baselines sit at 688.49
        // and 676.99 against 739.24 and 721.24.
        //
        // The grid does not re-snap afterwards either: `end` prints at 662.24,
        // which is one grid line below where the off-grid paragraph stopped
        // rather than the next multiple of the pitch — the same thing measured
        // after an `exact` paragraph.
        const ours = await linesOf('doc-grid-snap.docx');
        const printed = [757.24, 739.24, 721.24, 703.24, 688.49, 676.99, 662.24];

        expect(ours).toHaveLength(printed.length);
        printed.forEach((y, index) => {
            expect(Math.abs(ours[index]!.baselineY - y)).toBeLessThan(0.5);
        });
        expect(Math.round((ours[1]!.baselineY - ours[2]!.baselineY) * 100) / 100).toBe(18);
        expect(Math.round((ours[4]!.baselineY - ours[5]!.baselineY) * 100) / 100).toBe(11.5);
    });

    it('resolves a paragraph’s own spacing against the PITCH, except exact', async () => {
        // `doc-grid-spacing.docx` puts five paragraphs on an 18pt `w:docGrid`
        // over a font whose natural line is 11.50. Printed steps:
        //
        //   none          18.00   the pitch
        //   exact 14pt    14.00   itself — the one rule the grid does not touch
        //   atLeast 12pt  18.00   max(12, PITCH)
        //   auto 1.5      27.00   1.5 x PITCH, not 1.5 x the font's 11.50
        //   exact 24pt    24.00   itself
        //
        // Measuring the `exact` case ALONE reads as "a paragraph stating its
        // own w:spacing wins". It wins on HEIGHT, and only exact does. We
        // asked the font for all three and lost a third of the lines per page
        // under `atLeast`; that divergence was pinned at 28.24pt of drift by
        // the foot of this page, and it is nought now.
        const ours = await linesOf('doc-grid-spacing.docx');
        const stepOf = (first: number): number => Math.round(
            (ours[first]!.baselineY - ours[first + 1]!.baselineY) * 100) / 100;
        const printed = [
            757.24, 739.24, 721.24, 703.24, 687.24, 673.24, 657.24, 639.24,
            621.24, 603.24, 585.24, 558.24, 534.49, 513.49, 489.49, 468.49,
        ];

        expect(ours).toHaveLength(printed.length);
        expect([1, 4, 7, 10, 13].map(stepOf)).toEqual([18, 14, 18, 27, 24]);
        // Every baseline on the page, not only the steps: a rule that is right
        // about its own height and wrong about where it starts would pass the
        // line above and move everything under it.
        printed.forEach((y, index) => {
            expect(Math.abs(ours[index]!.baselineY - y)).toBeLessThan(0.5);
        });
    });

    it('centres an exact line in its own height ON a grid, where off one it uses the ratio', async () => {
        // The two rules are within half a point of each other at 12 and 14pt,
        // which is why the earlier print could not tell them apart. At 24 they
        // are 3.5 apart: the flat 0.8 ratio says 19.20 and centring the
        // natural line says (24 - 11.50) / 2 + 9.38 = 15.63. Printed, in a
        // paragraph placed FIRST so nothing unexplained sits above it: 15.62.
        //
        // Off a grid the ratio stands — it was verified in two fonts whose
        // natural baselines differ, and both printed the same 19.18 of a 24pt
        // line, which centring could not do.
        const ours = await linesOf('doc-grid-spacing-ends.docx');

        expect(Math.abs(ours[1]!.baselineY - 736.24)).toBeLessThan(0.5);
        expect(Math.round(ours[1]!.heightPx * 100) / 100).toBe(24);
        expect(Math.round((ours[0]!.baselineY - ours[1]!.baselineY) * 100) / 100).toBe(21);
    });

    it('gives back (multiple − 1) leadings at the FOOT of a gridded paragraph', async () => {
        // `doc-grid-spacing-ends.docx` runs the same 18pt grid over paragraphs
        // at 1.5 lines of one, three and five lines, and one at 2.0 lines.
        // Printed totals — measured gap-to-gap, so they include
        // whatever the paragraph does at its foot:
        //
        //   1.5, 1 line     23.75   = 1 x 27.00 - 3.25
        //   1.5, 3 lines    77.75   = 3 x 27.00 - 3.25
        //   1.5, 5 lines   131.75   = 5 x 27.00 - 3.25
        //   2.0, 3 lines   101.50   = 3 x 36.00 - 6.50
        //
        // So the constant is not a constant: it is (multiple - 1) times the
        // grid's own leading of 6.50, spent once at the foot. A single
        // paragraph at 1.5 would have said 3.25 and left it looking like the
        // half-leading, which is what it looked like at first.
        //
        // The plain and atLeast paragraphs of the fixture above give nothing
        // back — both total exactly N x 18.00 — which is the same rule at a
        // multiple of one.
        const ours = await linesOf('doc-grid-spacing-ends.docx');
        const printed = [
            757.24, 736.24, 712.24, 688.24, 667.24, 649.24, 625.49, 607.49,
            580.49, 553.49, 529.74, 511.74, 484.74, 457.74, 430.74, 403.74,
            379.99, 361.99, 325.99, 289.99, 260.49,
        ];

        expect(ours).toHaveLength(printed.length);
        printed.forEach((y, index) => {
            expect(Math.abs(ours[index]!.baselineY - y)).toBeLessThan(0.5);
        });
        // The last line of each 1.5-spaced paragraph is 3.25 short of the
        // 27.00 its neighbours take, and of the 2.0-spaced one 6.50 short.
        expect(Math.round(ours[5]!.heightPx * 100) / 100).toBe(23.75);
        expect(Math.round(ours[9]!.heightPx * 100) / 100).toBe(23.75);
        expect(Math.round(ours[19]!.heightPx * 100) / 100).toBe(29.5);
    });

    it('spends WHOLE grid lines on a line too tall for one', async () => {
        // `doc-grid-picture.docx`: an 18pt `w:docGrid` holding the same 18x36
        // picture. Printed — 757.24, then 715.89 for the picture's line, then
        // 685.24, 667.24 and 649.24, every one of those 18.00 apart.
        //
        // That puts the picture's line 54.00 deep, three pitches, with its
        // baseline 36.00 from the top: the line sits FLUSH with the top of the
        // group and the spare room falls below it, where a grid line that fits
        // its pitch is centred in it. Halving a negative leading put
        // our baseline 25.94 down a box of 18 — above its own box, over the
        // paragraph before it.
        const ours = await linesOf('doc-grid-picture.docx');
        const printed = [757.24, 715.89, 685.24, 667.24, 649.24];

        expect(ours).toHaveLength(printed.length);
        printed.forEach((y, index) => {
            expect(Math.abs(ours[index]!.baselineY - y)).toBeLessThan(0.5);
        });
        expect(Math.round(ours[1]!.heightPx * 10) / 10).toBe(54);
        // The grid's rhythm survives the tall line: every line after it is a
        // pitch apart, which is what "whole grid lines" buys.
        expect(Math.round((ours[2]!.baselineY - ours[3]!.baselineY) * 10) / 10).toBe(18);
    });
});

describe('a table cell is NOT on the section’s grid, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    /** The baselines of one cell, top to bottom, measured up from the foot. */
    const cellOf = async (name: string, rowIndex: number, cellIndex: number): Promise<number[]> => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return page.rows[rowIndex]!.cells[cellIndex]!.lines
            .map((line) => PAGE_HEIGHT_PT - pt(line.yPx + line.baselinePx));
    };

    it('steps a cell’s paragraphs by the FONT where the body steps by the pitch', async () => {
        // `onGrid` has skipped tables since the grid was first read, on the
        // stated grounds that "the grid is the SECTION's line, and a cell is
        // not on it" — asserted, never measured, and this arc has found that
        // shape of claim wrong more than once (the spacing claim above was
        // one).
        //
        // This time it is right. Printed on an 18pt grid: the body
        // paragraph above the table sits at 757.24 and the cell's three
        // paragraphs at 741.49, 729.99 and 718.49 — steps of 11.50, the font's
        // own line, where every body paragraph in the same section steps 18.00.
        const ours = await cellOf('doc-grid-cell.docx', 0, 0);

        expect(ours).toHaveLength(3);
        [741.49, 729.99, 718.49].forEach((y, index) => {
            expect(Math.abs(ours[index]! - y)).toBeLessThan(0.5);
        });
    });

    it('resolves a cell’s own multiple against the font, not the pitch', async () => {
        // The same table's second cell asks for 1.5 lines. On the body that
        // would be 27.00 — one and a half PITCHES — and in the cell
        // LibreOffice printed 741.49, 724.24, 706.99: steps of 17.25, which is
        // one and a half of the font's 11.50. A cell is off the grid for the
        // multiple as well as for the default.
        const ours = await cellOf('doc-grid-cell.docx', 0, 1);

        expect(ours).toHaveLength(3);
        [741.49, 724.24, 706.99].forEach((y, index) => {
            expect(Math.abs(ours[index]! - y)).toBeLessThan(0.5);
        });
    });

    it('gives up an atLeast floor in a GRIDDED cell, as the print does', async () => {
        // Once "small and strange", and pinned with OUR numbers for three
        // slices. A cell paragraph asking `atLeast` 12pt over an 11.50pt font
        // prints its lines 12.00 apart with no grid in the section — 694.74,
        // 682.74, 670.74, and `cell-spacing-nogrid` keeps proving it. Put a
        // grid on the section and LibreOffice drops the floor to 688.74,
        // 677.24, 665.74: steps of 11.50, LESS than the paragraph asked for
        // and nothing to do with the 18.00 pitch.
        //
        // Asking the same question four floors wide — in the body and in a
        // cell, with the grid and without — explained it: a cell under a grid
        // takes the font's own line and ignores the pitch AND the floor. That
        // is built, so these are now the PRINT's numbers rather than ours, and
        // the pair of fixtures is what keeps the two cases apart.
        const gridded = await cellOf('doc-grid-cell.docx', 1, 0);
        const plain = await cellOf('cell-spacing-nogrid.docx', 1, 0);

        [694.74, 682.74, 670.74].forEach((y, index) => {
            expect(Math.abs(plain[index]! - y)).toBeLessThan(0.5);
        });
        [688.74, 677.24, 665.74].forEach((y, index) => {
            expect(Math.abs(gridded[index]! - y)).toBeLessThan(0.5);
        });
    });
});

describe('the run children that break a line, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;
    const upFromFoot = (px: number): number => PAGE_HEIGHT_PT - pt(px);

    const drawnOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
            .map((op) => ({ text: op.text.trim(), xPx: pt(op.xPx), yPx: upFromFoot(op.yPx) }));
    };

    it('breaks the line at a w:cr, as it does at a w:br', async () => {
        // `w:cr` was in no branch of the run reader — not handled, not
        // reported, just passed over — so the text either side of one ran
        // together on a single line. Measured: `cc` printed at 714.49
        // and `dd` at 702.99, a step of 11.50, which is exactly what `aa` and
        // `bb` do either side of the `w:br` in the paragraph above.
        const drawn = await drawnOf('run-breaks.docx');
        const on = (text: string): number => drawn.find((op) => text === op.text)!.yPx;

        expect(Math.abs(on('aa') - 748.99)).toBeLessThan(0.5);
        expect(Math.abs(on('bb') - 737.49)).toBeLessThan(0.5);
        expect(Math.abs(on('cc') - 714.49)).toBeLessThan(0.5);
        expect(Math.abs(on('dd') - 702.99)).toBeLessThan(0.5);
    });

    it('chops a word too long for its line at the last character that FITS', async () => {
        // `word-chop.docx` puts two words in two measures each. Printed — the
        // same 16 letters, chopped where the width runs out:
        //
        //   gggggggghhhhhhhh in 51.3pt   ggggggggh · hhhhhhh
        //   gggggggghhhhhhhh in 31.3pt   ggggg · ggghh · hhhhh · h
        //   WWWWiiiiWWWWiiii in 51.3pt   WWWWiiii · WWWWiiii
        //   WWWWiiiiWWWWiiii in 31.3pt   WWW · WiiiiW · WWWi · iii
        //
        // The MIXED word is what proves the cut is fitted and not counted: the
        // measure that takes nine of the even word takes eight of it, and the
        // narrow one takes three, six, four and three — every line as many
        // letters as its own width allows.
        const drawn = await drawnOf('word-chop.docx');
        // Half a point either side, because our baselines sit that far from
        // the printed ones and a band cut to the exact numbers loses a line.
        const between = (top: number, foot: number): string[] => drawn
            .filter((op) => op.yPx <= top + 0.5 && op.yPx >= foot - 0.5)
            .map((op) => op.text);

        expect(between(748.99, 737.49)).toEqual(['ggggggggh', 'hhhhhhh']);
        expect(between(714.49, 679.99)).toEqual(['ggggg', 'ggghh', 'hhhhh', 'h']);
        expect(between(656.99, 645.49)).toEqual(['WWWWiiii', 'WWWWiiii']);
        expect(between(622.49, 587.99)).toEqual(['WWW', 'WiiiiW', 'WWWi', 'iii']);
    });

    it('declines the offer where the whole word fits, and draws nothing', async () => {
        // `soft-hyphen-fit.docx`, first paragraph: a measure wide enough for
        // the word. Printed: `eeeeeeee` at 72.10 and `ffffffff` at
        // 116.55 — one line, no hyphen, and 44.45 between them, which is the
        // eight letters alone. The character takes no width where the offer is
        // declined, which is why measuring it as the glyph most fonts keep at
        // U+00AD would push every such word 3.3pt wide.
        const drawn = await drawnOf('soft-hyphen-fit.docx');

        expect(drawn[1]!.text).toBe('eeeeeeee');
        expect(drawn[2]!.text).toBe('ffffffff');
        expect(Math.abs(drawn[1]!.yPx - drawn[2]!.yPx)).toBeLessThan(0.1);
        expect(Math.abs(drawn[2]!.xPx - 116.55)).toBeLessThan(0.5);
    });

    it('will not take an offer whose HYPHEN does not fit', async () => {
        // The same word in a 46pt measure: eight letters fit at 44.45 and the
        // hyphen they would owe does not. Printed: `eeeeeee` —
        // SEVEN — then `e` at 72.10 and `ffffffff` at 77.65 on the next line,
        // with no hyphen anywhere. LibreOffice declines the offer and chops one
        // letter short of it rather than draw a hyphen past the margin.
        //
        // We land on the same three, and only because the cost travels with
        // the text: the tail carries the offer, so it carries what accepting
        // costs, and it cannot rejoin the line its own head just left.
        const drawn = await drawnOf('soft-hyphen-fit.docx');
        const after = drawn.slice(4);

        expect(after.map((op) => op.text))
            .toEqual(['eeeeeee', 'e', 'ffffffff', 'gap2', 'eeeeeeee', 'ffffffff', 'end']);
        expect(Math.abs(after[0]!.yPx - 725.99)).toBeLessThan(0.5);
        expect(Math.abs(after[1]!.yPx - 714.49)).toBeLessThan(0.5);
        expect(Math.abs(after[1]!.xPx - 72.10)).toBeLessThan(0.5);
        expect(Math.abs(after[2]!.xPx - 77.65)).toBeLessThan(0.5);
        expect(drawn.some((op) => op.text.includes('-'))).toBe(false);
    });

    it('charges the hyphen only to a line that ENDS at the offer', async () => {
        // A 68pt measure: the word is 66.65 and the hyphen it would owe at the
        // offer is another 3.33. Printed: one line, `eeeeeeee` at
        // 72.10 and `ffffffff` at 116.55, no hyphen — carrying ON past an offer
        // costs nothing, and only ending there costs the hyphen.
        //
        // Both halves of that are needed and the fixtures now say so
        // separately: this measure breaks the moment the hyphen is charged to
        // the line that passes it, and the 46pt one above breaks the moment it
        // is not charged to the line that stops at it.
        const drawn = await drawnOf('soft-hyphen-fit.docx');
        const last = drawn.slice(-3, -1);

        expect(last.map((op) => op.text)).toEqual(['eeeeeeee', 'ffffffff']);
        expect(Math.abs(last[0]!.yPx - last[1]!.yPx)).toBeLessThan(0.1);
        expect(Math.abs(last[1]!.xPx - 116.55)).toBeLessThan(0.5);
    });

    it('breaks AT a w:softHyphen and draws the hyphen there', async () => {
        // A soft hyphen is an offer of a break and no ink until it is taken.
        // Printed: `eeeeeeee` then a `-` at x=116.55 ending that
        // line, with `ffffffff` beginning the next — where the same word
        // without one is chopped where the width runs out, `ggggggggh` /
        // `hhhhhhh`, and nothing is drawn.
        //
        // 116.55 is 44.45 past the margin at 72.10, which is exactly eight of
        // those letters: our own line begins at the same margin with the same
        // eight before the hyphen. This was pinned as a divergence first — the
        // character was in no branch of the reader — and chopping is what made
        // the FIT tractable: ending a line at the offer costs the hyphen,
        // carrying on past it costs nothing, and a segment already knew those
        // as two separate widths.
        const drawn = await drawnOf('run-breaks.docx');
        const offered = drawn.filter((op) => op.text.startsWith('eeee') || 'ffffffff' === op.text);
        const chopped = drawn.filter((op) => op.text.startsWith('gggg') || op.text.startsWith('hhhh'));

        expect(offered.map((op) => op.text)).toEqual(['eeeeeeee-', 'ffffffff']);
        expect(Math.abs(offered[0]!.yPx - 679.99)).toBeLessThan(0.5);
        expect(Math.abs(offered[1]!.yPx - 668.49)).toBeLessThan(0.5);
        // The control keeps its own break, so the hyphen is the difference and
        // not the narrower measure.
        expect(chopped.map((op) => op.text)).toEqual(['ggggggggh', 'hhhhhhh']);
    });
});

describe('what the reader used to pass over, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    const drawnOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
            .map((op) => ({
                text: op.text.trim(),
                xPx: pt(op.xPx),
                yPx: PAGE_HEIGHT_PT - pt(op.yPx),
            }));
    };

    it('drops a w:vanish run entirely — its ink AND its room', async () => {
        // Printed: `one ` at 72.10 and `two` at 91.55, with the
        // hidden run between them nowhere on the page and taking none of it —
        // 19.45 apart, which is the first word's own advance. Making it
        // invisible would leave the gap; the property means the text is not
        // there at all, and a reader that ignores it prints what was hidden.
        const drawn = await drawnOf('run-and-paragraph-properties.docx');

        expect(drawn.some((op) => op.text.includes('HIDDEN'))).toBe(false);
        expect(Math.abs(drawn[1]!.xPx - 72.10)).toBeLessThan(0.5);
        expect(Math.abs(drawn[2]!.xPx - 91.55)).toBeLessThan(0.5);
    });

    it('takes the LARGER of space-after and space-before, not their sum', async () => {
        // Ten points either side of a paragraph boundary: printed 21.50 apart,
        // which is the 11.50 line and ONE ten — not two. And the asymmetric
        // pairs settle which ten: 10 after then 20 before gives 31.50, and 20
        // after then 10 before gives 31.50 as well, so it is the larger of the
        // two and not one side of it.
        const drawn = await drawnOf('run-and-paragraph-properties.docx');
        const at = (text: string): number => drawn.find((op) => text === op.text)!.yPx;

        expect(Math.abs(at('plainA') - at('plainB') - 21.50)).toBeLessThan(0.5);
        expect(Math.abs(at('bigA') - at('bigB') - 31.50)).toBeLessThan(0.5);
        expect(Math.abs(at('smallA') - at('smallB') - 31.50)).toBeLessThan(0.5);
    });

    it('drops the space between paragraphs of one style for w:contextualSpacing', async () => {
        // Printed: the flagged pair sits 11.50 apart where the
        // identical pair without the flag sits 21.50, and the space goes
        // against a PLAIN neighbour of the same style as readily as against
        // another flagged one — so the flag belongs to the paragraph giving up
        // its own space rather than to the pair.
        //
        // Word's list styles set it as a matter of course. Until this was read
        // every list in every document was spaced like a run of ordinary
        // paragraphs.
        const drawn = await drawnOf('run-and-paragraph-properties.docx');
        const at = (text: string): number => drawn.find((op) => text === op.text)!.yPx;

        expect(Math.abs(at('ctxA') - at('ctxB') - 11.50)).toBeLessThan(0.5);
        expect(Math.abs(at('gap2') - at('ctxA') - 11.50)).toBeLessThan(0.5);
        // The unflagged pair, unchanged, in the same document.
        expect(Math.abs(at('plainA') - at('plainB') - 21.50)).toBeLessThan(0.5);
        // And the rule is SAME-STYLE, which needed a second paragraph style in
        // the fixture to see at all: `otherStyle` is `Probe` and spends nothing
        // below itself, so the only thing that could close the gap under it is
        // the flag — and the gap stands at 21.50. The identical pair with one
        // style either side closes to 11.50. Without both, a rule that dropped
        // the space against ANY neighbour passed every assertion above.
        expect(Math.abs(at('otherStyle') - at('ctxC') - 21.50)).toBeLessThan(0.5);
        expect(Math.abs(at('sameStyle') - at('ctxD') - 11.50)).toBeLessThan(0.5);
    });

    it('leaves a w:suppressLineNumbers paragraph OUT of the count', async () => {
        // Printed: `numbered` takes 10, `quiet` takes none, and `end` takes 11
        // — the line is not counted rather than counted and left blank, so
        // every number after it is one lower than it would otherwise be. The
        // whole page's numbering is asserted, because getting the skip right
        // and the count after it wrong is the failure this invites.
        const opened = await openWordFile(file('run-and-paragraph-properties.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const numbers = (page.lineNumbers ?? []).map((number) => number.run.text);

        expect(numbers).toEqual(Array.from({ length: 23 }, (_, index) => String(index + 1)));
        expect(numbers).toHaveLength(page.lines.length - 1);
    });
});

describe('w:pgNumType, end to end', () => {
    const textOfPage = (page: Parameters<typeof renderPage>[0]): string => renderPage(page).ops
        .filter((op): op is TextOp => 'text' === op.kind)
        .map((op) => op.text)
        .join('')
        .trim();

    it('numbers a section’s pages the way the SECTION asks, not always in digits', async () => {
        // `w:pgNumType/@w:fmt` is how a document asks for roman front matter,
        // and it is far commoner than a `PAGE \* roman` switch on each field —
        // which was all this engine read. Printed: `front i`,
        // `front ii`, then `body 1` where the next section restarts in digits.
        //
        // And the body's fields were not resolved AT ALL: a `PAGE` in a
        // paragraph printed whatever Word last cached, which in this fixture
        // is a literal `?`. They are answered now at the page they land on,
        // before the line is measured, because `i` and `iii` are not the same
        // width.
        const opened = await openWordFile(file('page-number-format.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pages.map(textOfPage)).toEqual(['front i', 'front ii', 'body 1']);
    });
});

describe('a page COUNT written in the body, end to end', () => {
    it('answers PAGE and NUMPAGES from the page each one lands on', async () => {
        // `page-count-in-body.docx` is three pages, each with `PAGE of
        // NUMPAGES` in its own text. Printed: `page 1 of 3`, `page 2
        // of 3`, `page 3 of 3`.
        //
        // Ours said `1 of 1`, `1 of 1`, `2 of 2` — two faults at once, both
        // uncovered by answering body fields at all. The page NUMBER was read
        // before the paragraph's own `pageBreakBefore` was taken, so
        // it named the page being left; and the COUNT was the pages laid so
        // far, because the second pass that exists to answer a count only ever
        // looked at headers and footers.
        const opened = await openWordFile(file('page-count-in-body.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);
        const textOf = (page: (typeof pages)[number]): string => renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => op.text)
            .join('')
            .trim();

        expect(pages.map(textOf)).toEqual(['page 1 of 3', 'page 2 of 3', 'page 3 of 3']);
    });
});

describe('w:tblStyle, end to end', () => {
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    /** Each table on the page, as the widths of the rules it draws. */
    const drawnBy = async () => {
        const opened = await openWordFile(file('table-style.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const tables = opened.document.paragraphs.filter(isTable);
        const rules = renderPage(page).ops.filter((op): op is LineOp => 'line' === op.kind);

        return { tables, rules, page };
    };

    it('takes a table’s borders from the STYLE it names', async () => {
        // Printed by LibreOffice: a table naming a style that draws
        // `w:sz="12"` and declaring nothing itself came out with rules 1.500
        // wide, all round and between. Nothing here read `w:tblStyle` at all,
        // so a table authored in Word — where "Table Grid" is the default and
        // the grid lives in the style — came out completely bare.
        const { tables } = await drawnBy();

        expect(tables).toHaveLength(6);
        expect(pt(tables[0]!.borders?.top?.widthPx ?? 0)).toBeCloseTo(1.5, 1);
        expect(pt(tables[0]!.borders?.insideH?.widthPx ?? 0)).toBeCloseTo(1.5, 1);
        // The control, which names no style: nothing to inherit and nothing drawn.
        expect(tables[3]!.borders?.top).toBeUndefined();
    });

    it('lets the table turn a side OFF that its style draws', async () => {
        // The same style, with all six sides declared `none` on the table:
        // LibreOffice printed no rule at all for it. So `none` is an answer
        // rather than a silence, and it beats the style.
        const { tables } = await drawnBy();

        expect(tables[1]!.borders?.top).toBeUndefined();
        expect(tables[1]!.borders?.insideH).toBeUndefined();
    });

    it('merges the two per SIDE, rather than one replacing the other', async () => {
        // The case that decided the shape of this. A table naming the same
        // 1.5pt style and declaring only its own `insideH` at 3pt printed the
        // middle rule at 3.000 and the four outer ones at 1.500 — so the
        // table's element does not replace the style's, and each side is
        // answered by the last level that names it.
        const { tables } = await drawnBy();

        expect(pt(tables[4]!.borders?.insideH?.widthPx ?? 0)).toBeCloseTo(3, 1);
        expect(pt(tables[4]!.borders?.top?.widthPx ?? 0)).toBeCloseTo(1.5, 1);
        expect(pt(tables[4]!.borders?.left?.widthPx ?? 0)).toBeCloseTo(1.5, 1);
    });

    it('lets a style BASED ON another override it, side by side', async () => {
        // How Word writes almost every table style. A leaf restating five
        // sides at half a point over a root drawing one and a half printed
        // exactly that — 0.500 round the outside and between the rows, and
        // 1.500 for the inside vertical the leaf leaves alone.
        const { tables } = await drawnBy();

        expect(pt(tables[5]!.borders?.top?.widthPx ?? 0)).toBeCloseTo(0.5, 1);
        expect(pt(tables[5]!.borders?.insideH?.widthPx ?? 0)).toBeCloseTo(0.5, 1);
        expect(pt(tables[5]!.borders?.insideV?.widthPx ?? 0)).toBeCloseTo(1.5, 1);
    });

    it('takes the CELL MARGINS from the style too', async () => {
        // Printed by LibreOffice: a style carrying `w:tblCellMar` of
        // 400 twips either side put the text at 91.85, where the table naming
        // no style at all put it at 77.50 — twenty points in from the page
        // margin rather than the format's default 5.4.
        const { page } = await drawnBy();
        const textOf = (tag: string): number => renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .filter((op) => op.text.trim().startsWith(tag))
            .map((op) => pt(op.xPx))[0]!;

        expect(textOf('C1')).toBeCloseTo(91.85, 0);
        expect(textOf('D1')).toBeCloseTo(77.50, 0);
    });
});

describe('what an inline VML box takes, at three widths', () => {
    // `text-box-inline-vml-widths.docx`, printed by LibreOffice. Shapes
    // stating 30, 90 and 180pt each took the stated width plus **18.0** of
    // the line, and their words sit **13.2** in from the box's left edge —
    // a constant, not a proportion, confirmed again beside a wrapping
    // paragraph.
    //
    // Two decimals is all the page supports: it reports to a hundredth and
    // this engine runs about 0.10 left of it, so the wrap probe's 18.01 and
    // the page's own 18.02 are one measurement, not two.
    //
    // That number was mis-read TWICE before it was measured: 17.50 first and
    // 16.40 next, both times because the box's start was estimated from the
    // width of the text before it instead of computed from the face.
    // `P-before` is 34.43, not the 36 that was assumed — a subtraction now
    // vendored as `tools/probes/run-width.mjs`. The page never moved.
    //
    // Still NOT built: this engine drops an inline VML box and says so. The
    // HEIGHT such a box gives its line was the missing number and is measured
    // in the block below.
    it('builds it, with the constant now accounted for', async () => {
        const opened = await openWordFile(file('text-box-inline-vml-widths.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const drawn = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => op.text.trim());

        // The words inside the boxes are the boxes' names, and each is now
        // drawn — where for a long time none was.
        expect(drawn).toContain('P');
        expect(drawn).toContain('S');
        expect(opened.document.diagnostics).toEqual([]);
    });
});

describe('what LibreOffice paints for an inline VML shape', () => {
    const pt = (px: number): number => px * 72 / 96;

    // `vml-shape-geometryless.docx`, printed by LibreOffice: a
    // `v:rect`, a `v:rect` holding text, and a bare `v:shape` — all 90x36pt,
    // the first and last carrying the SAME loud attributes (red fill, blue
    // stroke, 3pt of it).
    //
    // The print draws the rects and makes NO mark for the shape. That is a
    // controlled null rather than a quiet one: `pdf-strokes.py` finds the 3pt
    // stroke at width 3.000 on the rect, so a shape that draws nothing is the
    // renderer's answer and not a broken fixture.
    //
    // All three still put the run after them at the same 189.55, which is why
    // the fix drops the paint and keeps the room.
    it('paints the rects and nothing for the shape', async () => {
        const opened = await openWordFile(file('vml-shape-geometryless.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops;

        const fills = ops.filter((op) => 'rect' === op.kind).map((op) => op.fill);

        // Two rects, the colours they asked for, and no third box.
        expect(fills).toEqual(['#FF0000', '#00FF00']);

        // Specifically not the white-and-black default this engine used to
        // invent for a shape that names no colours it can honour.
        expect(fills).not.toContain('#FFFFFF');
    });

    it('reads the wrap distance the room is made of, and its default', async () => {
        // `vml-wrap-distance.docx`. Four 90pt rects: default, both
        // distances 0, both 20pt, and 20pt on the left alone. The print starts
        // the run after each at 189.55, 171.55, 211.55 and 191.55 — a box
        // beginning at 81.54, so 90 plus 18.0, 0, 40.0 and 20.0.
        //
        // This is the slice that turned the arc's oldest unexplained number
        // into a rule. It was read as 17.50, 16.40 and
        // 18.0 while nobody asked what it was MADE of; it is 9pt of
        // wrap distance a side, defaulted, and the file never says so because
        // the default is what it wants. The 20pt line is the control that
        // makes the 0 line evidence rather than silence.
        const opened = await openWordFile(file('vml-wrap-distance.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);

        const starts = ops.filter((op) => op.text.trim().startsWith('-')).map((op) => pt(op.xPx));
        const origin = pt(ops[0]!.xPx) + 9.44 + 90;

        expect(starts).toHaveLength(4);
        expect(starts[0]! - origin).toBeCloseTo(18.0, 0);
        expect(starts[1]! - origin).toBeCloseTo(0, 0);
        expect(starts[2]! - origin).toBeCloseTo(40.0, 0);
        expect(starts[3]! - origin).toBeCloseTo(20.0, 0);
    });

    it('gives every one of them the same room, whatever it holds', async () => {
        // `vml-inline-extra-room.docx`: six lines, each `L-` then a
        // 90pt shape then a marker. Default box, inset 0, inset 20pt, a 6pt
        // stroke, no stroke, and a shape with NO text box — and the print
        // starts all six markers at exactly 189.55.
        //
        // So the 18.0 of extra room belongs to the SHAPE and not to the text
        // box inside it, which is what says where it must be built. The inset
        // and stroke lines are NOT evidence: a control print moved a float's
        // text by nothing between inset 0 and inset 20pt, so that attribute
        // reaches nothing here and those two lines measured only their own
        // silence.
        const opened = await openWordFile(file('vml-inline-extra-room.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);

        const markers = ops.filter((op) => op.text.trim().startsWith('-')).map((op) => pt(op.xPx));

        expect(markers).toHaveLength(6);
        for (const x of markers) {
            // One x for all six, and it is the print's: the shape's 90, plus
            // 9pt of wrap distance either side.
            expect(x).toBeCloseTo(markers[0]!, 6);
            expect(pt(ops[0]!.xPx) + 9.44 + 90 + 18.0 - x).toBeCloseTo(0, 0);
        }
    });
});

describe('the inside of an inline VML box', () => {
    const pt = (px: number): number => px * 72 / 96;

    // `vml-box-inner-offset.docx` and `vml-box-content-width.docx`,
    // printed by LibreOffice. Everything a build of the box's CONTENT needs,
    // and one number explains all of it: an INSET of 4.25 on every side, on
    // top of the 9.0 of wrap distance either side.
    //
    // **Where the words start: 13.2 from the shape's origin.** Measured at
    // 90pt/10pt, 150pt/10pt and 90pt/20pt — 13.21, 13.21, 13.22. It follows
    // neither the font (double the size, same offset) nor the box width, so
    // it is furniture. 13.2 is the 9.0 of wrap plus the 4.25 of inset.
    //
    // **How wide the words may run: the stated width less 8.5** — the inset,
    // twice. Two rules survived every earlier probe, `stated - 13.2` (76.8 in
    // a 90pt box) and `stated - 8.5` (81.5), because every wrap measured so
    // far fell outside the gap between them. A string built to land IN it —
    // `alpha beta gammaw`, 79.97 by `run-width.mjs` — stayed on its line, so
    // 76.8 is refuted. With `alpha beta gamma delta` (94.68) wrapping at the
    // same width, and `alpha beta` (81.62 at 20pt) wrapping too, the content
    // width is bracketed to **[79.97, 81.62)**. 81.5 sits inside; 76.8 does not.
    //
    // **Where the first line sits: 4.25 below the box's top, plus the line's
    // own rise.** Box top less first baseline printed 13.55 at 10pt and 22.90
    // at 20pt, which look unrelated until the rise is taken off — this
    // engine's `lineHeight - descent`, 9.30 and 18.65. Both leave 4.25. The
    // inset arrived at from the side and from above, separately, and agreed.
    //
    // NOT built, deliberately: the words are still dropped and still said to
    // be. What stands between this and a build is the work of stacking blocks
    // into an `InlineShape`, which is already done for the DrawingML
    // spelling — not another number.
    it('puts the words where the print puts them, and wraps where it wraps', async () => {
        // Built in the slice after the one that measured it.
        const opened = await openWordFile(file('vml-box-content-width.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);
        const at = (text: string): TextOp => ops.find((op) => op.text.trim() === text)!;

        expect(opened.document.diagnostics).toEqual([]);

        // The 90pt box: `alpha beta gammaw` on one line and `zeta` under it,
        // which is the break that separated the two candidate content widths.
        expect(pt(at('gammaw').yPx)).toBeCloseTo(pt(at('alpha').yPx), 6);
        expect(pt(at('zeta').yPx)).toBeGreaterThan(pt(at('alpha').yPx));

        // 13.2 from the shape's origin — the line starts at 72.00 and `L-` is
        // 9.44 of it, so 81.44 plus the 9.0 of wrap and the 4.25 of inset.
        expect(pt(at('alpha').xPx) - (72 + 9.44)).toBeCloseTo(13.25, 0);

        // The first line sits its inset plus its own rise below the box's top,
        // and the box's foot stands on the outer baseline.
        const boxTopPt = pt(at('L-').yPx) - 60;

        expect(pt(at('alpha').yPx) - boxTopPt).toBeCloseTo(13.55, 0);
    });

    it('takes the inset off BOTH sides, which only the 20pt box can show', async () => {
        // `vml-box-inner-offset.docx`, third line: a 90pt box at 20pt text,
        // where the print puts `alpha`, `beta` and `gamma` each on a line of
        // their own — `alpha beta` is 81.62 and does not fit.
        //
        // This exists because a mutation charging the inset ONCE survived the
        // whole suite. 90 - 4.25 is 85.75, which is outside the measured
        // bracket of [79.97, 81.62) — but the only fixture that can tell 85.75
        // from 81.5 is this one, and nothing read it. The measurement was
        // vendored a slice before anything asserted it.
        const opened = await openWordFile(file('vml-box-inner-offset.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);

        // The 20pt box is the last of the three, so take the LAST of each word.
        const lastY = (text: string): number =>
            pt(ops.filter((op) => op.text.trim() === text).at(-1)!.yPx);

        expect(lastY('beta')).toBeGreaterThan(lastY('alpha'));
        expect(lastY('gamma')).toBeGreaterThan(lastY('beta'));
    });
});

describe('the height an inline VML box gives its line', () => {
    // `text-box-inline-vml-height.docx` and its `-carlito` twin, printed by
    // LibreOffice. Boxes stating 6, 12, 36 and 72pt, each with one
    // short word inside so the box's own height is the only thing varying,
    // between plain paragraphs that act as controls.
    //
    // **The rule: `max(natural line, box height + the font's DESCENT)`.** The
    // box's bottom sits ON the baseline — it takes the whole ascent side of
    // the line — and the line keeps its descent underneath. A box shorter than
    // the line it sits in disappears into it: the 6pt box stepped a plain
    // 11.55 at 10pt text and a plain 23.00 at 20pt.
    //
    // Measured, Liberation Serif, box height + extra:
    //   10pt text: natural 11.55, extra 2.25   (12 → 14.25, 36 → 38.25, 72 → 74.25)
    //   20pt text: natural 23.00, extra 4.35   (36 → 40.35, 72 → 76.35)
    //   40pt text: natural 46.00, extra 8.70   (60/90/120/150 all + 8.70)
    //
    // **Descent ALONE, not descent + line gap, and it took two fonts to say
    // so.** The extra scales with the text, which already ruled out box
    // furniture. Liberation Serif is what separates the two candidates: its
    // descent is 4.33 at 20pt and its descent + gap is 5.18, against a
    // measured 4.35. Carlito could never have shown it — its line gap is 0,
    // so both candidates predict its 5.37 against a measured 5.40. A control
    // font was not a formality here; it was the whole experiment.
    //
    // Both numbers come from the FACE (`vertical.descenderUnits`), not from a
    // fit: the same metrics predict the natural lines at 23.00 and 24.41
    // against 23.00 and 24.40 measured.
    //
    // One residual, honestly recorded: at 10pt the extra is 2.25 where the
    // face says 2.16, and the PLAIN line is 11.55 where the face says 11.50.
    // The anomaly is in the ordinary line too, so it is not a text-box
    // phenomenon and is not this rule's to explain.
    //
    // **This engine ALREADY does it, and the arc had it down as missing.**
    // An earlier pass recorded "the box's line stepped 38.25 where a plain
    // line steps 11.55" as an unexplained number blocking the build. It was neither
    // unexplained nor blocking: an inline VML shape is dropped only from the
    // FLOAT path, then falls through to the ordinary inline-shape path and
    // reserves its stated size — and this engine's inline-shape line rule is
    // already box-height-plus-descent. The steps below are the engine's, and
    // they are LibreOffice's to a fiftieth of a point.
    //
    // So the measurement stands and the conclusion drawn from it did not.
    // What it settles is CONFORMANCE — LibreOffice treats a VML text box's
    // line exactly as this engine treats any inline shape — which is worth
    // more than the missing rule it was mistaken for.
    //
    // The extra WIDTH is built now, the 18.0 having turned out to be 9pt of
    // wrap distance a side. What is still NOT built is the box's
    // CONTENT: the words inside it are dropped, and said to be.
    const pt = (px: number): number => px * 72 / 96;

    it.each([
        ['text-box-inline-vml-height.docx'],
        ['text-box-inline-vml-height-carlito.docx'],
    ])('builds it in %s, with the height rule on record', async (name) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const drawn = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => op.text.trim());

        // The controls print, and so does the word inside every box — one `x`
        // per box, four boxes.
        expect(drawn).toContain('p1');
        expect(drawn.filter((text) => 'x' === text)).toHaveLength(4);

        // And nothing is reported, because nothing is lost. This asserted a
        // drop for four slices; the message it asserted no longer exists.
        expect(opened.document.diagnostics).toEqual([]);
    });

    it.each([
        // Every step LibreOffice printed, in order: four plain lines, the 36pt
        // box, a plain line, the 72pt box, a plain line. The 6pt and 12pt
        // boxes are the two that vanish into the natural line.
        ['text-box-inline-vml-height.docx', 23.00, [23.00, 23.00, 23.00, 23.00, 40.35, 23.00, 76.35, 23.00]],
        ['text-box-inline-vml-height-carlito.docx', 24.40, [24.40, 24.40, 24.40, 24.40, 41.40, 24.40, 77.40, 24.40]],
    ])('steps through %s exactly as the print does', async (name, natural, printed) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const baselines = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            // The word INSIDE each box, now that it is built. These are the
            // outer paragraph's steps, which is what the print was read for.
            .filter((op) => 'x' !== op.text.trim())
            .map((op) => pt(op.yPx));

        // Downwards, because this engine's y grows down the page where the
        // PDF the rule was read off grows up it.
        const steps = baselines.slice(1).map((y, index) => y - baselines[index]!);

        expect(steps).toHaveLength(printed.length);
        steps.forEach((step, index) => expect(step).toBeCloseTo(printed[index]!, 0));

        // And the floor is a real floor, not an accident of these heights:
        // the two short boxes step the SAME as a line holding no box at all.
        expect(steps[1]).toBeCloseTo(natural, 0);
        expect(steps[3]).toBeCloseTo(natural, 0);
    });

    it('keeps the wrap distance either side of the box, as the print does', async () => {
        // `text-box-inline-vml-wrap.docx`: `L-`, a 90pt box, then `-R`. The
        // print starts `-R` at 189.55 — 81.54 in, plus the stated 90, plus the
        // 9pt of wrap distance either side.
        //
        // This was pinned as a SHORTFALL for two slices, asserting the gap was
        // NOT yet 18 wider precisely so it would fail on the day the constant
        // was built. It did.
        const opened = await openWordFile(file('text-box-inline-vml-wrap.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);

        const start = (text: string): number => pt(ops.find((op) => op.text.trim().startsWith(text))!.xPx);

        // `L-` starts the line; `-R` is 9.44 of `L-`, the box, and its wrap.
        expect(start('-') - start('L-')).toBeCloseTo(9.44 + 90 + 18.0, 0);
    });
});

describe('a table nested inside a TURNED cell', () => {
    // `cell-turned-nested-table.docx`, printed by LibreOffice: one
    // row, two cells holding the same three things — a paragraph, a bordered
    // one-cell table, another paragraph — with the left cell turned `btLr`.
    //
    // **The print DROPS the nested table's words in the turned cell.** Upright,
    // `Ubefore`, `Uin` and `Uafter` all come out. Turned, `Tbefore` and
    // `Tafter` come out on their sides — `[0 1 -1 0]`, a quarter turn — and
    // `Tin` appears nowhere on the page. It is not drawn small or off the
    // edge; `pdf-positions.py` lists the whole page and it is absent.
    //
    // **And this engine already does the same**, which the arc had carried as
    // an unbuilt pin rather than a conformance — the second time in this arc
    // that a standing "not built" turned out to be behaviour already in place.
    // Check the code before believing the pin.
    //
    // What DOES differ is the room the vanished table keeps. The print puts
    // `Tafter` 25.10 along the turned line from `Tbefore`; this engine puts it
    // 22.99 along. Both keep room for a thing neither draws, and ours is 2.11
    // short — pinned below, unexplained, and too small to guess at.
    it('drops the nested table in the turned cell and keeps it upright', async () => {
        const opened = await openWordFile(file('cell-turned-nested-table.docx'), FONTS);
        const pt = (px: number): number => px * 72 / 96;
        const at = new Map(layoutSections(opened.document.sections)
            .flatMap((page) => renderPage(page).ops)
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => [op.text.trim(), op]));

        // Turned: the paragraphs either side survive, the nested table's word
        // does not — exactly as the print has it.
        expect(at.has('Tbefore')).toBe(true);
        expect(at.has('Tafter')).toBe(true);
        expect(at.has('Tin')).toBe(false);

        // Upright, in the SAME row, all three. Without this the drop above
        // could be a nested table this reader never reads at all.
        expect(at.has('Ubefore')).toBe(true);
        expect(at.has('Uin')).toBe(true);
        expect(at.has('Uafter')).toBe(true);

        // The room kept for it, ours: 22.99 along the turned line, against the
        // print's 25.10. Asserted so the 2.11 cannot drift unnoticed.
        expect(pt(at.get('Tafter')!.xPx) - pt(at.get('Tbefore')!.xPx)).toBeCloseTo(22.99, 1);
    });
});

describe('the cell floor two prints cannot agree on', () => {
    // `grid-cell-floor.docx`, printed by LibreOffice: every
    // paragraph in a CELL, `atLeast` floors of 8, 12, 20 and 24 under an
    // 18.00 grid, then 12 and 20 under a 24.00 one, over 9pt text whose
    // natural line is 10.35.
    //
    // The print steps 18.00, 18.00, 20.00, 24.00 | 24.00, 24.00 — exactly
    // `max(floor, PITCH)`, tracking the pitch and not a number that happened
    // to be 18. Four floors and two pitches, and it looked settled.
    //
    // It is not. `doc-grid-cell.docx` has the SAME 18.00 pitch and the SAME
    // 12.00 floor and prints 11.50 — LibreOffice going UNDER the floor, which
    // stood unexplained through two passes.
    //
    // **It is not the font, and that claim stood here for two slices.** This
    // said "the one difference between the two files is the font". It is one
    // of at least three, and the deciding one is that `doc-grid-cell.docx`
    // carries a `word/settings.xml` and this hand-built file carries none.
    //
    // Ablated, one part at a time, reading p1 to p2:
    //   the file as it stands ............................. 11.50
    //   the SAME file with `word/settings.xml` removed .... 18.00  <- it snaps
    //   with Arial swapped for Liberation Serif ........... 11.55  <- font: no
    //   with only `<w:compat>` removed .................... 11.50  <- no
    //   with only `characterSpacingControl` removed ....... 11.50  <- no
    //   with settings.xml present but EMPTY ............... 11.50  <- !
    //
    // **There is no setting to find.** An empty `<w:settings/>` behaves like
    // the full one, so the trigger is the PART'S PRESENCE: LibreOffice gives a
    // document carrying a settings part one set of defaults and a document
    // without one another. The bisection set up to find a setting never had an
    // answer at the end of it.
    //
    // **Which turns this fixture into the suspect one.** `grid-cell-floor.docx`
    // is hand-built and carries no settings part, so its tidy
    // `max(floor, pitch)` — the rule nearly built from it — is the behaviour of
    // a file no real producer emits. Word and LibreOffice both always write
    // one. `doc-grid-cell-no-settings.docx` is vendored as the control that
    // says so: the real document, stripped, snapping to the grid it ignored.
    //
    // **And the rule, measured whole.** One document with a settings
    // part, floors of 8/12/20/24 over a font whose natural line is 11.55, the
    // same four pairs in the BODY and in a CELL:
    //
    //           floor:      8       12      20      24
    //   no grid   body:  11.55   12.00   20.00   24.00
    //   no grid   cell:  11.55   12.00   20.00   24.00
    //   grid 18   body:  18.00   18.00   20.00   24.00   <- max(floor, pitch)
    //   grid 18   cell:  11.55   11.55   11.55   11.55   <- neither
    //
    // Without a grid the floor is honoured everywhere, cells included, so it
    // was never "a cell ignores atLeast". WITH a grid, a cell ignores the grid
    // AND the floor and takes the font's own line — at 20 and 24, well clear
    // of 11.55, so this is not the floor quietly losing to something taller.
    // That is `doc-grid-cell.docx`'s 11.50 under a 12.00 floor, exactly.
    //
    // So that `max(floor, pitch)` was right about the BODY and read off a
    // cell, in a file whose missing settings part put LibreOffice in its
    // legacy defaults. Two wrong turns, one page.
    //
    // NOT built: acting on it means the reader knowing whether the package has
    // a settings part at all, which it has never needed to ask. OUR numbers
    // stay pinned until it does.
    it('keeps a cell’s own floor, and records what the print does instead', async () => {
        const opened = await openWordFile(file('grid-cell-floor.docx'), FONTS);
        const pt = (px: number): number => px * 72 / 96;
        const steps: number[] = [];

        for (const page of layoutSections(opened.document.sections)) {
            const at = new Map(renderPage(page).ops
                .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
                .map((op) => [op.text.trim(), pt(op.yPx)]));
            for (const name of ['A', 'B', 'C', 'D', 'E', 'F']) {
                const first = at.get(`${name}1`);
                const second = at.get(`${name}2`);
                if (undefined !== first && undefined !== second) {
                    steps.push(Math.round((second - first) * 100) / 100);
                }
            }
        }

        // The floors themselves, except where the font's own line is taller.
        expect(steps).toEqual([10.35, 12, 20, 24, 12, 20]);
    });

    it('gives every gridded cell the natural line, whatever floor it states', async () => {
        // `grid-cell-floor-modern.docx`: the table above, in a
        // document WITH a settings part and an 18pt grid. The print gives
        // every cell 11.55 — the font's own line — whatever floor it states.
        //
        // Pinned as a SHORTFALL when it was measured, so it would fail on the
        // day the rule was built. It failed one slice later, which
        // is what a shortfall pin is for.
        const opened = await openWordFile(file('grid-cell-floor-modern.docx'), FONTS);
        const pt = (px: number): number => px * 72 / 96;
        const at = new Map(layoutSections(opened.document.sections)
            .flatMap((page) => renderPage(page).ops)
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => [op.text.trim(), pt(op.yPx)]));

        const step = (name: string): number => at.get(`${name}b`)! - at.get(`${name}a`)!;

        // Every cell takes the natural line — at floors of 20 and 24, well
        // clear of 11.55, so this cannot be the floor quietly losing to a font
        // that happens to be tall.
        for (const name of ['C8', 'C12', 'C20', 'C24']) {
            expect(step(name)).toBeCloseTo(11.55, 0);
        }

        // And the BODY is untouched: `max(floor, pitch)`, which is the rule
        // that was found and put in the wrong place.
        expect(step('B8')).toBeCloseTo(18, 0);
        expect(step('B12')).toBeCloseTo(18, 0);
        expect(step('B20')).toBeCloseTo(20, 0);
        expect(step('B24')).toBeCloseTo(24, 0);
    });

    it('reads the ablated twin, so the pair cannot drift apart', async () => {
        // `doc-grid-cell-no-settings.docx` is `doc-grid-cell.docx` with its
        // `word/settings.xml` taken out and the part's declaration and
        // relationship taken out with it — an undeclared part prints EMPTY
        // rather than differently, which would read exactly like the
        // ablation having worked.
        //
        // Asserted here so the pair stays a pair: the twin must keep parsing,
        // and it must still be the same six paragraphs in the same two cells.
        // The 18.00 it PRINTS is not asserted, because this engine does not
        // build the snap yet — that is the open question above.
        const opened = await openWordFile(file('doc-grid-cell-no-settings.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const drawn = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => op.text.trim());

        // Arial is not in this repository's manifest, so the reader says it
        // substituted — and nothing else, which is the point: stripping the
        // part left a file with no other complaint.
        expect(opened.document.diagnostics.map((entry) => entry.kind)).toEqual(['font-substituted']);
        for (const name of ['p1', 'p2', 'p3', 'q1', 'q2', 'q3']) {
            expect(drawn).toContain(name);
        }
    });
});

describe('atLeast in a gridded section, and in a gridded CELL', () => {
    // `grid-at-least.docx`, printed by LibreOffice. An 18.00 grid
    // over 9pt text, with `w:lineRule="atLeast"` stated four ways: a floor
    // below the pitch and one above it, each in the body and in a cell.
    //
    // The reader believed `atLeast` needed nothing under a grid, on the
    // strength of a mutation that survived. It survived because nothing had
    // ever stated a floor ABOVE the pitch.
    const pt = (px: number): number => px * 72 / 96;

    const steps = async (): Promise<Record<string, number>> => {
        const opened = await openWordFile(file('grid-at-least.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const at = new Map(renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
            .map((op) => [op.text.trim(), pt(op.yPx)]));
        const step = (from: string, to: string): number =>
            Math.round(((at.get(to) ?? 0) - (at.get(from) ?? 0)) * 100) / 100;

        return {
            bodyUnder: step('A1', 'A2'),
            bodyOver: step('B1', 'B2'),
            cellUnder: step('C1', 'C2'),
            cellOver: step('D1', 'D2'),
        };
    };

    it('lets the pitch win where the floor is under it', async () => {
        // Both agree here, in the body and in the cell alike: 12.00 asked for,
        // 18.00 printed.
        const { bodyUnder } = await steps();

        expect(bodyUnder).toBe(18);
    });

    it('gives a floor ABOVE the pitch the whole of what it asked for', async () => {
        // BUILT. The print steps 24.00 and this engine stepped
        // 21.45, because a gridded paragraph gives leading back at its foot in
        // proportion to the MULTIPLE it asked for — `(24/18 - 1)` of the
        // grid's own leading, 2.55 — and a floor is not a multiple. 24.00 on
        // an 18.00 grid is not 1.33 of anything.
        //
        // The rule still says `grid`, because the baseline still sits in the
        // middle of its grid line: changing that instead moved every baseline
        // of `doc-grid-spacing.docx` by 3.26 and the older measurement caught
        // it. Only the giving back is skipped.
        const { bodyOver } = await steps();

        expect(bodyOver).toBe(24);
    });

    // ⚠️ Read this beside the print that measured a cell OFF the grid: its
    // paragraphs stepped the font's own 11.50 under an 18.00 pitch, and one
    // asking for 1.5 lines took 17.25 — one and a half of the FONT. The two
    // prints split the question rather than contradicting each other: a cell
    // is off the grid for the STEP it takes and, by the case below, on it for
    // the FLOOR a paragraph may not fall below. One printed case is a
    // hypothesis, so the next slice asks the same page for other floors and
    // for a pitch that is not 18.00 before writing a rule from it.
    it('MEASURED, NOT BUILT: a CELL is on the section’s pitch', async () => {
        // The other half, still pinned. The print steps 18.00 for a floor of
        // 12.00 inside a cell — the section's pitch, reaching in — where this
        // engine steps the 12.00 it was asked for, a cell being off the grid
        // for its own line rule. That a cell is off the grid in one sense and
        // on it in another is a rule about where the grid REACHES, and wants
        // its own slice.
        const { cellUnder, cellOver } = await steps();

        expect(cellUnder).toBe(12);
        // The case the two agree on, which is what says the difference is the
        // grid's reach and not the cell: a floor above the pitch steps its own
        // 24.00 in a cell as it does in the body.
        expect(cellOver).toBe(24);
    });
});

describe('a table inside an INLINE text box', () => {
    // `text-box-inline-table.docx`, printed by LibreOffice. The two
    // slices before this one met here and lost something between them: one
    // gave an inline box its own lines, the next let a box hold a table — and
    // an inline box carrying only lines dropped that table SILENTLY, since the
    // reader had stopped reporting what it no longer discards.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    it('draws the table inside the box, as the print does', async () => {
        // The print puts the two cells at 120.25 and 155.25 with the box's own
        // paragraph beneath them. Carried the way a float's rows are, which is
        // the same three lines in the other box's path.
        const opened = await openWordFile(file('text-box-inline-table.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const at = new Map(renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
            .map((op) => [op.text.trim(), pt(op.xPx)]));

        expect(at.get('T1')).toBeCloseTo(120.25, 0);
        expect(at.get('T2')).toBeCloseTo(155.25, 0);
        // And the words under it are still where they were measured.
        expect(at.get('INSIDE')).toBeCloseTo(115.35, 0);
    });
});

describe('a TABLE inside a text box', () => {
    // `text-box-table.docx`, printed by LibreOffice. A 2x2 table in a
    // VML box at 180pt, with a paragraph under it and body text beside it.
    // The reader dropped every such table and said so, because "a placed float
    // carries lines, and rows would want the renderer's row path as well as
    // its line path" — which turned out to be one call.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async () => {
        const opened = await openWordFile(file('text-box-table.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops;

        return {
            said: opened.document.diagnostics.map((entry) => entry.detail),
            cells: ops
                .filter((op): op is TextOp => 'text' === op.kind && op.text.startsWith('R'))
                .map((op) => pt(op.xPx)),
            verticals: [...new Set(ops
                .filter((op): op is LineOp => 'line' === op.kind
                    && Math.abs(op.x1Px - op.x2Px) < 0.01)
                .map((op) => pt(op.x1Px)))].sort((a, b) => a - b),
        };
    };

    it('lays the table out inside the box instead of dropping it', async () => {
        // The print draws the cells at 261.25 and 321.25 and the table's
        // verticals at 255.80, 315.80 and 375.80. This engine draws three
        // verticals in the same order 60.00 apart, and says nothing about
        // dropping anything.
        const { said, cells, verticals } = await drawn();

        expect(said).toEqual([]);
        expect(cells.length).toBe(4);
        expect(verticals.length).toBe(3);
        expect(verticals[1]! - verticals[0]!).toBeCloseTo(60, 1);
        expect(verticals[2]! - verticals[1]!).toBeCloseTo(60, 1);
    });

    it('puts it where this engine puts everything else in a VML box', async () => {
        // 258.70 against the print's 255.80, and the 2.90 is the VML inset's
        // pinned divergence rather than anything this slice did: LibreOffice
        // lays a VML box's content out on an inset of its own — 4.25, which it
        // uses
        // even for a box stating `inset="0,0,0,0"` — where this engine honours
        // the 0.1in the format states. The table lands exactly where the box's
        // own paragraphs land, which is the thing to hold on to.
        const { verticals } = await drawn();

        expect(verticals[0]).toBeCloseTo(258.70, 1);
    });
});

describe('an INLINE text box, which sits in the line rather than beside it', () => {
    // `text-box-inline.docx`, printed by LibreOffice. The reader has
    // carried a note for a long time — an inline text box "sits in the line
    // like a picture rather than beside the text, and no measurement covers
    // it, so it returns null and the caller says so". This is the measurement
    // that note was waiting for.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async () => {
        const opened = await openWordFile(file('text-box-inline.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);
        const at = new Map<string, number>();
        let previous: { text: string; at: number } | null = null;

        for (const op of ops) {
            const text = op.text.trim();
            if ('' === text) {
                continue;
            }
            if (text.endsWith('-')) {
                previous = { text, at: pt(op.xPx) };
            } else if (null !== previous) {
                at.set(previous.text + text, previous.at);
                previous = null;
            } else {
                at.set(text, pt(op.xPx));
            }
        }

        const baselines = new Map<string, number>();
        for (const op of ops) {
            if ('' !== op.text.trim()) {
                baselines.set(op.text.trim().split(' ')[0]!, pt(op.yPx));
            }
        }

        return {
            at,
            lines: (word: string): number => baselines.get(word) ?? 0,
            said: opened.document.diagnostics.map((entry) => entry.detail),
        };
    };

    it('gives the box its extent and draws the words inside it', async () => {
        // BUILT on the measurement above. A 90x36pt inline box takes
        // exactly its stated `wp:extent`: the print draws its word at 115.35 —
        // the box's own left edge plus the 7.20 body inset — and the text
        // after it at 198.20, which is 90.00 past where the box begins.
        //
        // It is a SHAPE with words: the piece already measured its own width
        // and grew its line, and only the words needed somewhere to live. They
        // are stacked when the box is READ, which a float's cannot be — an
        // inline box's inner width is settled by its extent before anything
        // knows where the box lands.
        // Its words WRAP at the box's own inner width — 90.00 less the 7.20
        // inset on either side — which is the only thing that says they were
        // stacked to the box rather than to the page: the print breaks
        // `INSIDE THE BOX` after `THE`, putting both lines at 115.35 and the
        // second 11.55 under the first.
        const { at, lines } = await drawn();

        expect(at.get('INSIDE')).toBeCloseTo(115.35, 0);
        expect(at.get('BOX')).toBeCloseTo(115.35, 0);
        expect(lines('BOX') - lines('INSIDE')).toBeCloseTo(11.55, 0);
        expect(at.get('A-after')).toBeCloseTo(198.20, 0);
    });

    it('leaves the VML spelling for its own measurement, and says so', async () => {
        // The WORDS are still not built, and are still said to be dropped.
        //
        // The room is. This once read "nothing in the file
        // accounts for the extra 17.50" — and something did: the shape's wrap
        // distance, 9pt a side by default, which the file leaves unstated
        // because that IS the default. So the run after it moved 18.0 right,
        // from 197.54 to 215.54, and that is the constant being built rather
        // than a regression.
        //
        // It was reaching the float reader once — a VML shape with
        // no `position:absolute` is INLINE — and came out as a float with no
        // offsets at all, placed at the paragraph's origin with its words
        // drawn nowhere and NOTHING said about it. Now it is dropped and said,
        // and the drawn-shape path keeps the box's 90.00 plus its wrap.
        const { at, said } = await drawn();

        expect(at.get('B-after')).toBeCloseTo(215.54, 1);
        // Its words are drawn now — 9.0 of wrap and 4.25 of inset
        // in from where the box's room begins — and nothing is reported,
        // because nothing is lost.
        expect(at.has('VMLIN')).toBe(true);
        expect(said).toEqual([]);
    });
});

describe('w:object — an embedded thing, and the picture Word keeps of it', () => {
    // `embedded-object.docx`, printed by LibreOffice. A chart, an
    // equation, a spreadsheet: the object itself is not something this engine
    // can run, and Word stores a VML picture of what it looks like beside it.
    // The reader has handled that since forever and no fixture had one, so the
    // branch had never drawn anything.
    //
    // Three rows of `before` + something + `after`, so the x of `after` says
    // what room the something took.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async () => {
        const opened = await openWordFile(file('embedded-object.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops;
        const words = new Map<string, number>();
        let previous: { text: string; at: number } | null = null;

        for (const op of ops) {
            if ('text' === op.kind && '' !== op.text.trim()) {
                // The renderer splits a line at its spaces, so `A-after`
                // arrives as `A-` then `after` — and the x that matters is
                // where the FIRST of them starts, which is where the print
                // begins the whole word.
                if (op.text.trim().endsWith('-')) {
                    previous = { text: op.text.trim(), at: pt(op.xPx) };
                } else if (null !== previous) {
                    words.set(previous.text + op.text.trim(), previous.at);
                    previous = null;
                }
            }
        }

        return {
            words,
            images: ops
                .filter((op): op is ImageOp => 'image' === op.kind)
                .map((op) => ({ at: pt(op.xPx), width: pt(op.widthPx), height: pt(op.heightPx) })),
        };
    };

    it('draws the object’s picture exactly where a bare picture would go', async () => {
        // The print puts the object's picture at 108.05 by 60x30 with the text
        // after it at 168.20, and the same VML picture OUTSIDE any object at
        // 107.50 with its text at 167.65. The two are one answer: an embedded
        // object is its picture, in the place a picture takes.
        const { words, images } = await drawn();

        expect(images[0]).toEqual({ at: 108.09, width: 60, height: 30 });
        expect(images[1]).toEqual({ at: 107.54, width: 60, height: 30 });
        expect(words.get('A-after')).toBeCloseTo(168.20, 0);
        expect(words.get('B-after')).toBeCloseTo(167.65, 0);
    });

    it('keeps NO room for an object whose picture is missing, where the print keeps 78.11', async () => {
        // MEASURED, NOT BUILT. A `w:object` whose shape carries no
        // `v:imagedata` has nothing to draw, and both renderers agree there is
        // nothing to draw — the print's page holds no image, no rule and no
        // fill for that row. It reserves the space anyway: its text lands at
        // 185.65 where this engine, dropping the object outright, puts it at
        // 107.54.
        //
        // 78.11pt, and nothing in the file says 78.11 — not the shape's stated
        // 60x30, nor the object's own `w:dxaOrig` of 1200 twips, which is the
        // same 60. A blank of a size the document does not state is not
        // something to reproduce by guessing at it, so the number is recorded
        // and the gap left open.
        const { words, images } = await drawn();

        expect(images.length).toBe(2);
        expect(words.get('C-after')).toBeCloseTo(107.54, 1);
    });
});

describe('the width of a string that is not Latin', () => {
    // `cyrillic-widths.docx`, printed by LibreOffice. Until now no
    // non-Latin string's WIDTH had ever been put beside a printed page.
    //
    // `real-lease.docx` is in Russian, which is why that gap could go unseen:
    // every line of it starts at the left margin and none of them wraps, so a
    // Cyrillic glyph resolving to `.notdef` — or to the wrong glyph — would
    // move nothing in the one fixture that carries the language.
    //
    // Right-aligned, so the x of a row IS the right margin minus its width.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async (): Promise<Map<string, number>> => {
        const opened = await openWordFile(file('cyrillic-widths.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return new Map(renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
            .map((op) => [op.text.trim(), pt(op.xPx)]));
    };

    it('measures Cyrillic and Greek as the print measures them', async () => {
        // The three rows that are one run each, so their x is directly
        // comparable. Every one lands 0.17 left of the print, which after its
        // own 0.10 frame is 0.07pt of width across eight glyphs — under a
        // hundredth of a point each.
        const at = await drawn();

        expect(at.get('ЖЩЪЫЬЭЮЯ')).toBeCloseTo(459.35, 0);
        expect(at.get('ийклмнопр')).toBeCloseTo(475.90, 0);
        expect(at.get('ΑΒΓΔΕΖΗΘ')).toBeCloseTo(470.70, 0);
    });

    it('starts a mixed row where the print starts it', async () => {
        // `Договор Dogovor` in one run, so the two scripts are measured
        // together: 449.80 printed, and the Latin word after it at 487.75.
        // A row of Cyrillic and a row of Latin the same length apart — 454.30
        // against 458.20 — is what says the two are not being measured alike
        // by accident.
        const at = await drawn();

        expect(at.get('Договор')).toBeCloseTo(449.80, 0);
        expect(at.get('Dogovor')).toBeCloseTo(487.75, 0);
    });

    it('finds every one of those code points in the face', async () => {
        // The failure this guards against does not announce itself: a missing
        // glyph is drawn as `.notdef`, which HAS a width, so the line still
        // measures and simply comes out wrong — the same trap as a hyphen the
        // font lacked.
        const face = FONTS.resolve('Liberation Serif', false, false).font;
        const missing = [...'Договор аренды ЖЩЪЫЬЭЮЯ ийклмнопр ΑΒΓΔΕΖΗΘ']
            .filter((character) => 0 === face.glyphId(character.codePointAt(0)!));

        expect(missing).toEqual([]);
    });
});

describe('the footnote separator a document tries to change', () => {
    // `footnote-separator-emptied.docx` — `footnotes.docx` with the paragraph
    // inside its `w:separator` footnote emptied, which is what an author does
    // by deleting the separator line in Word. Printed by LibreOffice.
    //
    // `w:separator` and `w:continuationSeparator` have been on the unread list
    // since the survey began, in five fixtures and named nowhere in `src`.
    // This settles them.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 1000) / 1000;

    const ruleOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const rule = renderPage(page).ops.find((op): op is LineOp => 'line' === op.kind)!;

        return {
            from: pt(rule.x1Px),
            to: pt(rule.x2Px),
            width: pt(rule.widthPx),
            notes: page.footnotes!.lines
                .map((line) => line.line.pieces.map((piece) => piece.text).join('')),
        };
    };

    it('draws its own rule whatever the file says, exactly as the print does', async () => {
        // The separator is the RENDERER's furniture, not the document's
        // content. Three states of the same file print the same rule at
        // 28.350 to 172.350 by 0.100: the standard `<w:separator/>`, the
        // paragraph emptied, and both separator footnotes removed outright.
        //
        // So leaving the two elements unread is right, and the engine's own
        // rule — two inches at the margin, a hairline, 43.949 above the foot —
        // is what a reader sees however the file is edited.
        const standard = await ruleOf('footnotes.docx');
        const emptied = await ruleOf('footnote-separator-emptied.docx');

        expect([emptied.from, emptied.to, emptied.width]).toEqual([28.35, 172.35, 0.1]);
        expect([emptied.from, emptied.to, emptied.width])
            .toEqual([standard.from, standard.to, standard.width]);
    });

    it('leaves the notes themselves alone', async () => {
        // The control: emptying the separator must not disturb what it
        // separates, which is the way a reader of the file could have broken
        // this while trying to honour it.
        const emptied = await ruleOf('footnote-separator-emptied.docx');

        expect(emptied.notes).toEqual(['1 first note']);
    });
});

describe('widows, orphans, and the one default the reference disagrees about', () => {
    // `widow-orphan.docx`, printed by LibreOffice. A page holding
    // six lines, filled to a chosen depth and then given a three-line
    // paragraph — built with explicit breaks, so the count is exact whatever
    // the text measures — so the question is which of its lines page one keeps.
    const pagesOf = async (): Promise<string[][]> => {
        const opened = await openWordFile(file('widow-orphan.docx'), FONTS);

        return layoutSections(opened.document.sections).map((page) => page.lines
            .map((line) => line.line.pieces.map((piece) => piece.text).join(''))
            .filter((text) => '' !== text));
    };

    const wherever = async (text: string): Promise<number> =>
        (await pagesOf()).findIndex((page) => page.includes(text));

    it('keeps two lines together at both ends of a paragraph', async () => {
        // Five fillers leave room for one line: this engine moves the whole
        // paragraph rather than strand its first. Four leave room for two, and
        // it moves a second line across rather than strand the last.
        const pages = await pagesOf();

        expect(await wherever('A-line1')).toBe(await wherever('A-line2'));
        expect(pages[await wherever('B-line3')]).toContain('B-line2');
    });

    it('is the ONE default LibreOffice reads the other way', async () => {
        // MEASURED AND DELIBERATELY DIFFERENT. With the element absent the
        // print allowed both: a page ended `A-fill5, A-line1` with lines two
        // and three overleaf, and another ended `B-line1, B-line2` with the
        // third alone on the next page.
        //
        // Kept as it is because the thing being modelled is WORD's layout —
        // the format states the default is true and Word applies it — while
        // LibreOffice is the reference this engine can measure rather than the
        // authority on what a `.docx` means. Not one document in this corpus
        // states the element, including the Word-authored one, so the default
        // decides the pagination of all of them.
        expect(await wherever('A-line1')).not.toBe(await wherever('A-fill5'));
    });

    it('honours the element where a document states it, as the print does', async () => {
        // The divergence is confined to SILENCE. Asked for outright, both
        // renderers move the whole paragraph: the print put C's three lines on
        // a page of their own and D's likewise, and so does this.
        const pages = await pagesOf();

        expect(pages[await wherever('C-line1')]).toEqual(['C-line1', 'C-line2', 'C-line3']);
        expect(pages[await wherever('D-line1')]).toEqual(['D-line1', 'D-line2', 'D-line3']);
    });

    it('moves a w:keepLines paragraph whole, which both agree on', async () => {
        // E is three lines with two of them fitting, and both renderers move
        // all three — but so does the widow rule on its own, so E cannot tell
        // the two apart. G is FOUR lines with two fitting: widow and orphan
        // have nothing to say about it, and only `w:keepLines` moves it. The
        // print left its page with the fillers alone.
        const pages = await pagesOf();

        expect(pages[await wherever('E-line1')]).toEqual(['E-line1', 'E-line2', 'E-line3']);
        expect(pages[await wherever('G-line1')])
            .toEqual(['G-line1', 'G-line2', 'G-line3', 'G-line4']);
    });

    it('lets ONE paragraph turn the rule off while the document keeps it on', async () => {
        // The other half of "per paragraph": F states `w:widowControl="0"`
        // against a document default of on, and the print left its third line
        // alone overleaf — the widow allowed, for that paragraph only. It is
        // the one case that can tell a paragraph's own answer from the
        // document's, since every other section here agrees with the default.
        const pages = await pagesOf();

        expect(pages[await wherever('F-line2')]).toContain('F-fill4');
        expect(pages[await wherever('F-line3')]).toEqual(['F-line3']);
    });
});

describe('a DOUBLE strike and a double underline', () => {
    // `decoration-double.docx`, printed by LibreOffice. `w:dstrike`
    // was in no fixture at all and a double underline had never had its gap
    // printed — both were drawn a whole thickness apart because that reads as
    // two lines, which is a choice and not a measurement.
    //
    // Each decoration at 10pt and 40pt, so the answer says whether the gap
    // scales with the font.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const rulesUnder = async (name: string): Promise<{ at: number; width: number }[]> => {
        const opened = await openWordFile(file('decoration-double.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops;
        const line = ops.find((op): op is TextOp => 'text' === op.kind && op.text.startsWith(name))!;

        // Within half the run's own size of its baseline: a decoration sits a
        // fraction of the em away, and a fixed window in points reaches into
        // the next paragraph at 10pt while missing nothing at 40.
        return ops
            .filter((op): op is LineOp => 'line' === op.kind
                && Math.abs(op.y1Px - line.yPx) < line.sizePx / 2)
            .map((op) => ({ at: pt(line.yPx - op.y1Px), width: pt(op.widthPx) }));
    };

    it('draws the pair a whole thickness apart, where the print thins and straddles', async () => {
        // MEASURED, NOT MATCHED. The print struck 10pt text at +3.00 and +1.70
        // with a 0.40 rule where its single is +2.60 at 0.60 — thinner than
        // the single, and straddling it. This engine keeps the single's
        // thickness and puts both rules to one side of it.
        //
        // Not matched because the SINGLE is not: this engine takes its
        // underline and strikeout from the FONT, where LibreOffice uses
        // something of its own that six measurements have not identified
        // (`TrueTypeFont.decoration`). A double built on somebody else's
        // single would be neither one thing nor the other.
        const strike = await rulesUnder('C');

        expect(strike.map((entry) => entry.at)).toEqual([2.05, 3.03]);
        expect(strike.map((entry) => entry.width)).toEqual([0.49, 0.49]);
    });

    it('scales the pair with the font, as the print does', async () => {
        // The one thing both renderers agree on. At 40pt the print's pair is
        // 4.20 apart (11.80 and 7.60) where its 10pt pair is 1.30; ours is
        // 3.91 against 0.98. Four times the size, four times the gap, on both
        // sides — so the gap is a multiple of the thickness in both, and only
        // the multiple differs.
        const small = await rulesUnder('G');
        const large = await rulesUnder('H');
        const gap = (rules: { at: number }[]): number =>
            Math.round(Math.abs(rules[1]!.at - rules[0]!.at) * 100) / 100;

        expect(gap(small)).toBeCloseTo(0.98, 1);
        expect(gap(large)).toBeCloseTo(3.91, 1);
        expect(gap(large) / gap(small)).toBeCloseTo(4, 1);
    });
});

describe('w:lvlRestart, which the reference renderer ignores', () => {
    // `list-restart.docx`, printed by LibreOffice. Three lists of the
    // same three-level shape — `1. a) i. ii. b) ? 2. a) ?` — differing only in
    // what the third level says about restarting: nothing, `0`, and `1`.
    const markersOf = async (list: string): Promise<string[]> => {
        const opened = await openWordFile(file('list-restart.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return page.lines
            .filter((line) => line.line.pieces.map((piece) => piece.text).join('')
                .startsWith(list))
            .map((line) => line.marker?.run.text ?? '-');
    };

    it('counts straight through a level that says it never restarts', async () => {
        // MEASURED AND DELIBERATELY DIFFERENT. The print numbers all three
        // lists identically — `1. a) i. ii. b) i. 2. a) i.` — so LibreOffice
        // ignores `w:lvlRestart` outright, in every spelling.
        //
        // We follow the FILE, as we do with the VML inset, and for a
        // second reason: Word honours the element — it is the "restart
        // numbering after" setting — so a document authored there reads
        // correctly this way and not the other. It is how a document numbers
        // its figures 1..40 across chapters that each restart their sections.
        expect(await markersOf('B')).toEqual(
            ['1.', 'a)', 'i.', 'ii.', 'b)', 'iii.', '2.', 'a)', 'iv.']);
    });

    it('restarts under any shallower level where nothing forbids it', async () => {
        // The default, and the one case where this engine and the print agree:
        // both restart the roman under each new letter and each new number.
        expect(await markersOf('A')).toEqual(
            ['1.', 'a)', 'i.', 'ii.', 'b)', 'i.', '2.', 'a)', 'i.']);
    });

    it('says so where the stated restart LEVEL is not the one it applies', async () => {
        // `w:lvlRestart="1"` names the level whose change restarts this one —
        // level one, so a new letter should NOT restart the roman while a new
        // number should. This engine restarts under any shallower level, which
        // happens to print what LibreOffice printed and would not print what
        // Word prints.
        //
        // Unmeasurable against a renderer that ignores the attribute, so it is
        // reported rather than left silent.
        const opened = await openWordFile(file('list-restart.docx'), FONTS);
        const said = opened.document.diagnostics
            .filter((entry) => entry.detail.includes('w:lvlRestart'))
            .map((entry) => entry.detail);

        expect(await markersOf('C')).toEqual(await markersOf('A'));
        expect(said).toEqual([
            'w:lvlRestart "1" is read as restarting under any level above',
        ]);
    });
});

describe('columns of unequal width, and the toggle that decides them', () => {
    // `columns-unequal.docx`, printed by LibreOffice. `columns.docx`
    // states `<w:cols w:num="2" w:space="708"/>` and nothing else, so the
    // equal-width arithmetic was measured and the OTHER branch never had been.
    //
    // Four sections, each two columns stating 3000 twips, a 500 gap and 5526 —
    // which fills the 9026 of writing width exactly — and each spelling the
    // toggle differently. The page is short on purpose: a column holds six
    // lines, so eight paragraphs reach the second one.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const secondColumn = async (): Promise<Map<string, number>> => {
        const opened = await openWordFile(file('columns-unequal.docx'), FONTS);
        const found = new Map<string, number>();

        for (const page of layoutSections(opened.document.sections)) {
            for (const op of renderPage(page).ops) {
                // The seventh paragraph of each section is the first to reach
                // the second column, and its name says which section it is.
                if ('text' === op.kind && op.text.endsWith('7')) {
                    found.set(op.text, pt(op.xPx));
                }
            }
        }

        return found;
    };

    it('uses a stated width unless the document asks for equal ones', async () => {
        // 247.10 is 72 + 150 + 25: the first column's own width and the gap it
        // states. Divided evenly with the default half-inch gap it would be
        // 315.75, which is where the section that says `w:equalWidth="1"` puts
        // it — and the whole of the difference between the two branches.
        const at = await secondColumn();

        expect(at.get('A7')).toBeCloseTo(247.10, 0);
        expect(at.get('B7')).toBeCloseTo(247.10, 0);
        expect(at.get('C7')).toBeCloseTo(247.10, 0);
        expect(at.get('D7')).toBeCloseTo(315.75, 0);
    });

    it('reads the toggle in every spelling, and an ABSENT one as off', async () => {
        // The schema says `w:equalWidth` defaults to TRUE. The page disagrees:
        // section B says nothing at all and still gets its stated widths. So
        // the default decides only the case where no width is stated, which
        // this branch is never asked about.
        //
        // Reading it as `"0"` alone — which is what Word writes, and all this
        // engine accepted — got A and D right and B and C wrong.
        const at = await secondColumn();

        expect(at.get('B7')).toBe(at.get('A7'));
        expect(at.get('C7')).toBe(at.get('A7'));
        expect(at.get('D7')).not.toBe(at.get('A7'));
    });
});

describe('a cell that spans columns, and where a bordered table hangs', () => {
    // `cell-grid-span.docx`, printed by LibreOffice. `w:gridSpan` is
    // read, carried across the editor seam, and was in NO vendored fixture —
    // `vertical-merge.docx` covers the other axis and nothing covered this
    // one, though a merged header row is what half the tables in the world
    // look like.
    //
    // Three columns of 100pt. Row 1 spans 1-2, row 3 spans all three, row 4
    // spans 2-3, and row 2 is the plain control. A second table below states a
    // 3pt border where the first states 1pt.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async () => {
        const opened = await openWordFile(file('cell-grid-span.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops;

        const verticals = ops.filter((op): op is LineOp => 'line' === op.kind
            && Math.abs(op.x1Px - op.x2Px) < 0.01);

        return {
            at: new Map(ops
                .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
                .map((op) => [op.text.trim(), pt(op.xPx)])),
            verticals: verticals.map((op) => pt(op.x1Px)),
            /**
             * Which rows the rule down column `x` runs beside.
             *
             * Asked of the ROWS rather than as a length, because the two
             * renderers mitre a corner differently — the print runs its
             * verticals half a horizontal rule past each end, 51.20 down the
             * left edge against our 50.00 — and that is a question about
             * corners, not about spans.
             */
            beside: (x: number): number[] => page.rows
                .map((row, index) => ({ index, middle: row.yPx + row.heightPx / 2 }))
                .filter(({ middle }) => verticals.some((op) => Math.abs(pt(op.x1Px) - x) < 0.1
                    && Math.min(op.y1Px, op.y2Px) <= middle
                    && Math.max(op.y1Px, op.y2Px) >= middle))
                .map(({ index }) => index),
        };
    };

    it('gives a spanning cell the whole width of the columns it covers', async () => {
        // The cell after a span starts where the GRID says, not one column on:
        // `C3` printed at 277.00 beside a cell covering columns one and two,
        // exactly where `C3b` sits in the unspanned row beneath it.
        const { at } = await drawn();

        expect(at.get('SPAN12')).toBeCloseTo(77.00, 0);
        expect(at.get('C3')).toBeCloseTo(277.00, 0);
        expect(at.get('C3b')).toBeCloseTo(at.get('C3')!, 1);
        expect(at.get('SPANALL')).toBeCloseTo(77.00, 0);
        expect(at.get('SPAN23')).toBeCloseTo(177.00, 0);
    });

    it('draws no rule where a span swallowed the boundary', async () => {
        // The half a merged cell can get wrong without moving any text.
        //
        // Row indices run across all three tables: 0-3 are the spanned one,
        // 4 is the thick-bordered one, 5-6 the one with mixed borders.
        //
        // The print runs the first table's left edge past all four of its
        // rows, its first inside boundary only beside rows two and four —
        // 757.289 to 743.789 and 732.189 to 719.689 — and its second only
        // beside rows one and two, 768.889 to 743.789. Every spanned cell is a
        // gap in a rule that would otherwise be continuous.
        const { beside } = await drawn();

        expect(beside(171.5)).toEqual([1, 3, 5, 6]);
        expect(beside(271.5)).toEqual([0, 1, 5, 6]);
        // The left edge has no span to interrupt it, in any table that hangs
        // at 71.50 — which is every one whose first row states a 1pt rule.
        expect(beside(71.5)).toEqual([0, 1, 2, 3, 5, 6]);
        expect(beside(70.5)).toEqual([4]);
    });

    it('hangs the table by half the rule its CELLS declare', async () => {
        // The defect this fixture found. A table hangs its left border into the
        // margin — measured at two widths already — but the width was read off
        // the TABLE's own border, and a table authored in Word carries its
        // borders on the cells.
        //
        // Printed with the border stated only on the cells: the 1pt table's
        // left rule centred on 71.50 and the 3pt table's on 70.50, both putting
        // the rule's outer edge on the 72pt margin. Reading only the table's
        // border left each of them half a rule too far right.
        const { verticals, at } = await drawn();

        expect(verticals.some((x) => Math.abs(x - 71.50) < 0.1)).toBe(true);
        expect(verticals.some((x) => Math.abs(x - 70.50) < 0.1)).toBe(true);
        // And the text moves with it: a whole point between the two tables,
        // which is the difference between half of 1pt and half of 3pt.
        expect(at.get('SPAN12')! - at.get('THICK')!).toBeCloseTo(1.00, 1);
    });

    it('runs its verticals half a rule PAST the corner, not up to it', async () => {
        // The print puts this table's horizontal rules at 769.389 and
        // 719.189 and runs its left edge from 769.889 to 718.689 — half a rule
        // beyond each, so the corner is filled. Stopping on the centre line,
        // which is what this engine did, leaves a quarter of every outer
        // corner unpainted.
        //
        // Asserted as the OVERHANG rather than as the two absolute numbers,
        // because the table's own height is a hundredth or two adrift of the
        // print's and that is a different question.
        const opened = await openWordFile(file('cell-grid-span.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const ops = renderPage(page).ops.filter((op): op is LineOp => 'line' === op.kind);
        const topRule = Math.min(...ops
            .filter((op) => Math.abs(op.y1Px - op.y2Px) < 0.01)
            .map((op) => op.y1Px));
        const leftEdge = Math.min(...ops
            .filter((op) => Math.abs(op.x1Px - op.x2Px) < 0.01)
            .map((op) => Math.min(op.y1Px, op.y2Px)));

        expect(pt(topRule - leftEdge)).toBeCloseTo(0.50, 1);
    });

    it('takes the hang from the FIRST row, not the widest rule in the table', async () => {
        // Where guessing by analogy went wrong. `ruleAbove` resolves a shared
        // horizontal rule widest-wins, so the left edge looked like it should
        // too — but a table whose first row claims 1pt and whose second claims
        // 3pt printed BOTH rules centred on 71.50, the first row's hang, with
        // the thicker one spilling either side of that same edge.
        //
        // Its text stays at 77.00 in both rows, where a hang of 1.50 would put
        // it at 76.00 as the all-3pt table's is.
        const { at, verticals } = await drawn();

        expect(at.get('MIXTHIN')).toBeCloseTo(77.00, 0);
        expect(at.get('MIXTHICK')).toBeCloseTo(77.00, 0);
        expect(at.get('MIXTHICK')).toBeCloseTo(at.get('SPAN12')!, 1);
        // The 3pt rule is drawn, and drawn on the same line as the 1pt one.
        expect(verticals.filter((x) => Math.abs(x - 71.50) < 0.1).length)
            .toBeGreaterThan(4);
    });
});

describe('revision marks, and the view this engine takes of them', () => {
    // `revision-marks.docx`, printed by LibreOffice. Five paragraphs
    // of `before` + something + `after`, so the x of `after` reports what room
    // the something took: a tracked deletion, a tracked insertion, a
    // hyperlink, the abandoned half of a tracked move, and nothing at all.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const linesOf = async () => {
        const opened = await openWordFile(file('revision-marks.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return {
            text: page.lines.map((line) => line.line.pieces.map((piece) => piece.text).join('')),
            rules: renderPage(page).ops.filter((op): op is LineOp => 'line' === op.kind),
            said: opened.document.diagnostics.map((entry) => entry.detail),
        };
    };

    it('draws the document as it would READ, not as an editor sees it', async () => {
        // MEASURED, NOT BUILT. LibreOffice prints tracked changes as MARKUP:
        // `DELETED` drawn in place and struck through by a 0.60pt rule 44.90
        // long, 2.60 above the baseline; `INSERTED` underlined by one 48.30
        // long, 1.20 below it; the moved-from run struck TWICE, at 3.00 and
        // 1.70; and a change bar down the margin at 64.95, which is 7.15pt
        // left of the text.
        //
        // That is the view a word processor gives an editor. This engine draws
        // what the document says once the changes are accepted — the reading
        // view, which is what a preview of an uploaded file is for — so a
        // deletion takes NO room and no rule is drawn for any of it.
        const { text, rules } = await linesOf();

        expect(text[0]).toBe('before  after');
        expect(text[3]).toBe('before  after');
        expect(text[0]).toBe(text[4]);
        expect(rules.length).toBe(0);
    });

    it('keeps what an insertion and a hyperlink wrap', async () => {
        // The other half of the same decision: a `w:ins` is accepted text and
        // a `w:hyperlink` only wraps runs, so both are drawn — and both agree
        // with the print, which puts them at 100.10 like everything else here.
        const { text } = await linesOf();

        expect(text[1]).toBe('before INSERTED after');
        expect(text[2]).toBe('before LINKED after');
    });

    it('says that it dropped something, once', async () => {
        // Dropping text silently is the one thing a reader must not do — the
        // same rule that covers a character it cannot draw. Two
        // paragraphs lose text here and the notice is one, because the detail
        // is the same and `report` keeps only the first of those.
        const { said } = await linesOf();

        expect(said).toEqual(['text deleted under revision tracking is not drawn']);
    });
});

describe('an ALIGNED stop under an indent, and a tab in a FOOTER', () => {
    // `tab-aligned.docx`, printed by LibreOffice. Two combinations the corpus
    // had none of, both downstream of the indent work. `tab-alignment.docx`
    // has no indent in it, so a centre, right or decimal stop had only ever
    // been measured in a paragraph starting at the margin — and an aligned
    // stop places the text AFTER it, so its arithmetic runs through the very
    // origin that slice changed. No header or footer anywhere in the corpus
    // contained a tab at all, which is the `left | centre | right` footer
    // every document in the world has.
    //
    // Every body row aims at one stop, 4320 twips from the margin: page 288.10.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async (): Promise<Map<string, number>> => {
        const opened = await openWordFile(file('tab-aligned.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const found = new Map<string, number>();

        for (const op of renderPage(page).ops) {
            if ('text' === op.kind && '' !== op.text.trim()) {
                found.set(op.text.trim(), pt(op.xPx));
            }
        }

        return found;
    };

    it('aligns against the COLUMN’s stop, whatever the paragraph’s indent', async () => {
        // `MMMM` centred on the stop printed at 270.35 with no indent, and at
        // the same 270.35 under a 25pt one. A stop belonging to the paragraph
        // would have centred the second at 295.35.
        const at = await drawn();

        expect(at.get('A')).toBeCloseTo(72.10, 0);
        expect(at.get('B')).toBeCloseTo(97.10, 0);
        // Both rows draw the same word, so the map holds the LAST — which is
        // the indented one, and its equality with the print is the point.
        expect(at.get('MMMM')).toBeCloseTo(252.55, 0);
    });

    it('ends a right-aligned run on the stop and puts a decimal separator ON it', async () => {
        // Right: `MMMM` is 35.56pt wide and printed at 252.55, ending at
        // 288.11. Decimal: the print splits `12.34` into `12` at 278.10 and
        // `.34` at 288.10 — the SEPARATOR sits on the stop, not the string.
        const at = await drawn();
        const opened = await openWordFile(file('tab-aligned.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const decimal = renderPage(page).ops
            .find((op): op is TextOp => 'text' === op.kind && op.text.includes('12.34'))!;
        const beforeSeparator = decimal.font.measureUnkerned('12', decimal.sizePx);

        expect(pt(decimal.xPx + beforeSeparator)).toBeCloseTo(288.10, 0);
        expect(at.get('12.34')).toBeCloseTo(278.10, 0);
    });

    it('collapses a tab whose text cannot fit before the stop', async () => {
        // Forty `M`s are 355.60pt and the stop is 216 along, so there is no
        // room to range them against it. The print gives the tab NO width at
        // all — the text follows the `E` directly, at 103.25 — rather than
        // pulling it back to before the line's own start.
        const at = await drawn();

        expect(at.get('MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM')).toBeCloseTo(103.25, 0);
    });

    it('gives a FOOTER its stops, and the column’s defaults when it declares none', async () => {
        // The three-column footer: a centre stop at half the text width and a
        // right stop at the right margin put MIDDLE at 278.35 and RIGHT at
        // 492.85. The paragraph under it declares no stops at all and falls to
        // the defaults, 36pt apart from the margin — so a footer has no
        // implicit centre or right stop of its own, whatever Word's own
        // template puts in its style.
        const at = await drawn();

        expect(at.get('LEFT')).toBeCloseTo(72.10, 0);
        expect(at.get('MIDDLE')).toBeCloseTo(278.35, 0);
        expect(at.get('RIGHT')).toBeCloseTo(492.85, 0);

        expect(at.get('bare')).toBeCloseTo(72.10, 0);
        expect(at.get('none')).toBeCloseTo(108.10, 0);
        expect(at.get('stated')).toBeCloseTo(144.10, 0);
    });
});

describe('a whole contract, against the page LibreOffice printed', () => {
    /**
     * AUTHORED, not found, and the reason is measurable: of the fifteen
     * documents this platform has ever generated, not one contains a tab, a
     * `w:ind` or a contents list. So the rules measured for markers and for
     * indents — a list marker and the tab after it, a stop inside an indented
     * paragraph, an aligned stop with a leader — had never met in one file,
     * and no real file in this repo could make them meet.
     *
     * It is the weaker kind of whole-document guard, since its author knew
     * what it was for. It still earned its place on the first print: the
     * signature rule came out one underscore short.
     */
    const PAGE_HEIGHT_PT = 841.89; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    /** Where each line begins, top of the page down, as LibreOffice printed it. */
    const PRINTED_LEFT = [
        227.20, 72.10, 72.10, 72.10, 72.10, 72.10, 72.10, 100.45, 99.60, 99.00,
        72.10, 100.45, 99.60, 86.30, 100.45, 72.10, 72.10, 72.00,
    ];

    it('starts every line where the print starts it', async () => {
        const opened = await openWordFile(file('contract.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const rows = new Map<number, number>();

        for (const op of renderPage(page).ops) {
            if ('text' !== op.kind || '' === op.text) {
                continue;
            }

            const y = Math.round((PAGE_HEIGHT_PT - pt(op.yPx)) * 100) / 100;
            rows.set(y, Math.min(rows.get(y) ?? Infinity, pt(op.xPx)));
        }

        const ours = [...rows.entries()].sort((a, b) => b[0] - a[0]).map(([, x]) => x);

        expect(ours.length).toBe(PRINTED_LEFT.length);
        expect(Math.max(...ours.map((x, row) => Math.abs(x - PRINTED_LEFT[row]!))))
            .toBeLessThan(0.5);
    });

    it('lands every tab where the print lands it', async () => {
        // Four kinds in one file: a right-aligned stop at the right margin
        // carrying a dot leader (the contents), a left stop halfway across
        // (the signature columns), a tab inside an INDENTED paragraph, and a
        // tab on the second line of a hanging one. The last two are the pair
        // no real document in this repo could have exercised.
        const opened = await openWordFile(file('contract.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const at = (text: string): number[] => renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && op.text.trim() === text)
            .map((op) => Math.round(pt(op.xPx) * 100) / 100);

        for (const landed of at('1')) {
            expect(landed).toBeCloseTo(517.90, 0);
        }
        for (const landed of at('(see')) {
            expect(landed).toBeCloseTo(396.10, 0);
        }

        expect(at('Interest')[0]).toBeCloseTo(252.10, 0);
        expect(at('Tenant')[0]).toBeCloseTo(297.75, 0);
    });

    it('draws the marker of a RIGHT-justified list level by its own width', async () => {
        // The sub-clauses number `(a)`, `(b)` at a level whose `w:lvlJc` is
        // right, so each marker ends on the anchor and its own width decides
        // where it starts: `(a)` printed at 99.60 and the wider `(b)` at 99.00.
        // A whole document is where that rule and the tab after it meet.
        const opened = await openWordFile(file('contract.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const markers = page.lines
            .filter((line) => undefined !== line.marker && line.marker.run.text.startsWith('('))
            .map((line) => Math.round(pt(line.marker!.xPx) * 100) / 100);

        expect(markers[0]).toBeCloseTo(99.60, 0);
        expect(markers[1]).toBeCloseTo(99.00, 0);
        expect(markers[0]).not.toBe(markers[1]);
    });
});

describe('how many glyphs a tab leader draws', () => {
    // `tab-leader-fill.docx`, printed by LibreOffice. Each row is an
    // empty run and one tab to a right-aligned stop, so the span the leader
    // fills IS the stop's position and the arithmetic is exact. The spans are
    // chosen to land just over, halfway, just under and EXACTLY on a glyph —
    // only the exact one tells `ceil` from "one more than fits" — and one
    // under a single glyph, which turned out to matter most.
    const drawn = async (): Promise<{ glyph: string; count: number }[]> => {
        const opened = await openWordFile(file('tab-leader-fill.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && /^[._\-·]+$/.test(op.text))
            .map((op) => ({ glyph: op.text[0]!, count: op.text.length }));
    };

    it('never overshoots with dots', async () => {
        // 140.109, 140.509, 140.909 and 140.982 of them fit, and the print
        // drew 140 every time — so not `round`, which would have given 141 for
        // three of those. An exact 141.000 drew 141.
        const runs = await drawn();
        const dots = runs.filter((run) => '.' === run.glyph).map((run) => run.count);

        expect(dots).toEqual([140, 140, 140, 141, 140]);
    });

    it('always passes the stop with a rule, by exactly one glyph', async () => {
        // The finding. A span of EXACTLY 36 underscores drew 37, which is what
        // rules out `ceil` and leaves "one more than fits" — a rule to sign on
        // reaches its stop rather than stopping 5.50pt short of it.
        //
        // The last of these is a span half a glyph wide, which drew its one
        // underscore anyway.
        const runs = await drawn();
        const rules = runs.filter((run) => '_' === run.glyph).map((run) => run.count);

        expect(rules).toEqual([37, 37, 37, 37, 36, 37, 1]);
    });

    it('goes to the nearer for a hyphen or a middle dot, and draws none under one', async () => {
        // 54.804 drew 55 and 54.053 drew 54 — nearest, for both glyphs, which
        // have the same advance. But 0.546 drew NOTHING, where rounding alone
        // would have made it one: not one whole glyph, not one glyph.
        const runs = await drawn();

        expect(runs.filter((run) => '-' === run.glyph).map((run) => run.count)).toEqual([55, 54]);
        expect(runs.filter((run) => '·' === run.glyph).map((run) => run.count))
            .toEqual([55, 54]);
    });

    it('draws no dot leader where a whole dot does not fit', async () => {
        // 0.727 of a dot printed nothing, which floor already gives — the case
        // is here because the same span drew one UNDERSCORE, and a rule that
        // treated all leaders alike would have to be wrong about one of them.
        const runs = await drawn();

        expect(runs.filter((run) => '.' === run.glyph).length).toBe(5);
    });
});

describe('what a tab stop is measured FROM', () => {
    // `tab-origin.docx`, printed by LibreOffice. Each row is `a`, a
    // tab, then its own name, so the x of the name is where the tab landed.
    // The indents are chosen so that measuring from the margin and measuring
    // from the paragraph's indent cannot give the same answer:
    //
    //   P1  no indent, stop at 1600           152.10   the control: both agree
    //   P2  indent 720, stop at 1600          152.10   from the indent: 188.10
    //   P3  indent 500, no explicit stop      108.10   from the indent: 133.10
    //   P4  indent 720 + firstLine 360,       152.10   from the indent: 206.10
    //       stop at 1600
    //   P5  indent 500, stop at 600           102.10   from the indent: 127.10
    //   P6  indent 500, stop at 300           108.10   a stop the indent puts
    //                                                  BEHIND the line's start
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const landed = async (): Promise<Map<string, number>> => {
        const opened = await openWordFile(file('tab-origin.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const found = new Map<string, number>();

        for (const op of renderPage(page).ops) {
            if ('text' === op.kind && op.text.startsWith('P')) {
                found.set(op.text, pt(op.xPx));
            }
        }

        return found;
    };

    it('measures explicit stops from the MARGIN, not from the paragraph', async () => {
        // The same stop at 80pt, once in an unindented paragraph and once
        // under a 36pt indent, lands in the same column. Measured from the
        // paragraph it would sit 36pt further on, which is where every tab in
        // an indented paragraph used to go.
        const at = await landed();

        expect(at.get('P1')).toBeCloseTo(152.10, 0);
        expect(at.get('P2')).toBeCloseTo(152.10, 0);
        expect(at.get('P2')).toBe(at.get('P1'));
    });

    it('measures the DEFAULT stops from there too', async () => {
        // A 25pt indent and no explicit stop: the tab goes to 36pt, the first
        // default stop in the column, and not to 25 + 36.
        const at = await landed();

        expect(at.get('P3')).toBeCloseTo(108.10, 0);
    });

    it('does not let a first-line indent shift them either', async () => {
        // P4's first line starts 18pt further in than its own paragraph, and
        // its tab still lands in the column's 80pt stop. This half was already
        // right — the origin allowed for the first line but not for the indent
        // under it — which is why the two are asserted apart.
        //
        // P7 is the same paragraph with a line BREAK before its tab, so the
        // tabbed text is on line two. That line starts back at the plain
        // indent, 108.10, and its tab lands at the same 152.10: the first-line
        // indent counts on the first line and nowhere else. Without a second
        // line in the fixture, an origin that charged the first-line indent to
        // every line looked exactly like one that did not.
        const at = await landed();

        expect(at.get('P4')).toBeCloseTo(152.10, 0);
        expect(at.get('P7')).toBeCloseTo(152.10, 0);
    });

    it('skips a stop the indent puts behind the line, and falls to the defaults', async () => {
        // P5's stop at 30pt is past the 25pt indent, so it is reachable and the
        // tab barely moves — 102.10, just clear of the `a`. P6's at 15pt is
        // BEHIND where its line even starts, so nothing can land on it and the
        // default stops take over at 36pt. That is the same rule the breaker
        // already followed past the last explicit stop; this is the case where
        // the last explicit stop is also the first.
        const at = await landed();

        expect(at.get('P5')).toBeCloseTo(102.10, 0);
        expect(at.get('P6')).toBeCloseTo(108.10, 0);
    });
});

describe('w:kern, which decides whether a string is kerned at all', () => {
    // `kerning.docx`, printed by LibreOffice. Five RIGHT-aligned
    // paragraphs, so the x of each one's first glyph is the right margin minus
    // its width: a difference between two rows IS a difference in measured
    // width, and no frame has to be lined up.
    //
    //   A  AVAVAVAVAV, nothing said     451.25   nine strong kern pairs
    //   B  AVAVAVAVAV, w:kern 16        462.85   kern above 8pt: the text is 10pt
    //   C  AVAVAVAVAV, w:kern 40        462.85   kern above 20pt: the text is 10pt
    //   D  xxxxxxxxxx, nothing said     473.45   no kern pair in the string
    //   E  xxxxxxxxxx, w:kern 16        473.45
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawn = async (): Promise<TextOp[]> => {
        const opened = await openWordFile(file('kerning.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return renderPage(page).ops.filter((op): op is TextOp => 'text' === op.kind);
    };

    it('does not kern a run that never asked to be kerned', async () => {
        // The finding, and the whole reason this fixture exists: the row with
        // no `w:kern` printed 11.60pt WIDER than the row with one. The wider
        // number is the plain sum of advances — so a page kerns only where the
        // document says so, and every string here was being kerned.
        const ops = await drawn();

        expect(pt(ops[1]!.xPx) - pt(ops[0]!.xPx)).toBeCloseTo(462.85 - 451.25, 0);
        expect(ops[0]!.kerned).toBe(undefined);
        expect(ops[1]!.kerned).toBe(true);
    });

    it('reads w:kern as a switch and ignores the size it states', async () => {
        // `w:kern="40"` asks for kerning above 20pt over 10pt text, so by the
        // format's own reading it should do nothing. The print puts it at the
        // same 462.85 as `w:kern="16"`, kerned — LibreOffice reads only whether
        // the element is THERE. Honouring the threshold would put us at odds
        // with the page for exactly the files that state one.
        const ops = await drawn();

        expect(pt(ops[2]!.xPx)).toBe(pt(ops[1]!.xPx));
        expect(ops[2]!.kerned).toBe(true);
    });

    it('leaves a string with no kern PAIR where it was, either way', async () => {
        // The control that says the 11.60 above is kerning and not something
        // the element does to the paragraph: `xxxxxxxxxx` has no pair to kern,
        // and its two rows printed at one x whether `w:kern` was there or not.
        const ops = await drawn();

        expect(pt(ops[4]!.xPx)).toBe(pt(ops[3]!.xPx));
        expect(ops[3]!.kerned).toBe(undefined);
        expect(ops[4]!.kerned).toBe(true);
    });

    it('tells the SVG which, rather than leaving it to the viewer', async () => {
        // A renderer draws a piece as a string and the drawing engine steps
        // between the glyphs itself — kerning by default, in every browser. An
        // unkerned run has to say so or its later glyphs land left of where
        // the line breaker put them.
        const opened = await openWordFile(file('kerning.docx'), FONTS);
        const svg = renderPageToSvg(layoutSections(opened.document.sections)[0]!);

        expect(svg).toContain('font-kerning:none');
        expect(svg).toContain('font-kerning:normal');
    });
});

describe('where a list marker sits, and where the text after it starts', () => {
    // `list-marker-justified.docx`, printed by LibreOffice. Eleven
    // lists of three items each, numbered from 9 so a one-digit and a two-digit
    // marker stand side by side — that difference is what reports the
    // justification. Every item's text names its own row.
    //
    //   A-C  hanging 360, roomy:      left / right / center
    //   D-F  hanging 120, too small:  left / right / center
    //   G    like A, but the paragraph states no `w:ind` of its own
    //   H    indent 500, hanging 120  — an indent that is NOT a multiple of the
    //                                   default tab, so a stop measured from the
    //                                   margin and one measured from the indent
    //                                   give different answers
    //   I    like D, plus an explicit `w:tabs` stop at 1600
    //   J    indent 500, hanging 360  — the marker ends BEFORE the indent
    //   K    level says 720/360, the PARAGRAPH says 1000/200
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    /** Each item's row, by the text in it: where its marker and its text begin. */
    const rows = async (): Promise<Map<string, { marker: number; text: number }>> => {
        const opened = await openWordFile(file('list-marker-justified.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const found = new Map<string, { marker: number; text: number }>();

        for (const line of page.lines) {
            found.set(line.line.pieces.map((piece) => piece.text).join(''), {
                marker: pt(line.marker?.xPx ?? 0),
                text: pt(line.xPx),
            });
        }

        return found;
    };

    it('justifies the marker around one ANCHOR — it does not fill the slot', async () => {
        // All three justifications work off `indentLeft - hanging`, which here
        // is 90.10. `left` starts there; `right` ENDS there, growing away from
        // the text; `center` straddles it. Read off the print: a right-aligned
        // `10.` begins at 77.60 and a centred one at 83.85, against the 90.10
        // where every left-aligned marker begins whatever its width.
        const at = await rows();

        expect(at.get('a9')!.marker).toBeCloseTo(90.10, 0);
        expect(at.get('a10')!.marker).toBeCloseTo(90.10, 0);
        expect(at.get('b9')!.marker).toBeCloseTo(82.60, 0);
        expect(at.get('b10')!.marker).toBeCloseTo(77.60, 0);
        expect(at.get('c9')!.marker).toBeCloseTo(86.35, 0);
        expect(at.get('c10')!.marker).toBeCloseTo(83.85, 0);
    });

    it('sends the text to the first tab stop past the marker, not snugly clear of it', async () => {
        // Group D's marker overruns its 6pt hanging space by 6.50pt. A snug
        // push would start the text at 114.60; the print has 144.10 — the next
        // default stop. The rule is the SUFFIX: `w:suff` is a tab by default,
        // and a tab goes to a stop.
        //
        // Group F is the same rule caught mid-step: a centred `9.` ends before
        // the indent and its text starts at 108.10, while `10.` — wider by one
        // digit — ends past it and its text jumps the whole way to 144.10.
        const at = await rows();

        expect(at.get('d9')!.text).toBeCloseTo(144.10, 0);
        expect(at.get('f9')!.text).toBeCloseTo(108.10, 0);
        expect(at.get('f10')!.text).toBeCloseTo(144.10, 0);
        // A right-justified marker grows away from the text, so it never
        // pushes: group E has the same 6pt hanging space and stays at 108.10.
        expect(at.get('e9')!.text).toBeCloseTo(108.10, 0);
        expect(at.get('e10')!.text).toBeCloseTo(108.10, 0);
    });

    it('counts the indent as a stop, and explicit stops before the defaults', async () => {
        // J: the marker ends at 91.60 and the indent is 97.10 — the text starts
        // AT the indent, so the indent is itself a stop.
        //
        // H: the same 25pt indent with a marker that runs past it. The text
        // goes to 108.10, the first default stop measured FROM THE MARGIN. A
        // stop measured from the indent would be 133.10, which is the number
        // this file was built to be able to disagree with.
        //
        // I: an explicit stop at 80pt takes the text to 152.10 — past the
        // default stop at 144.10 that would otherwise have caught it.
        const at = await rows();

        expect(at.get('j9')!.text).toBeCloseTo(97.10, 0);
        expect(at.get('h9')!.text).toBeCloseTo(108.10, 0);
        expect(at.get('h10')!.text).toBeCloseTo(108.10, 0);
        expect(at.get('i9')!.text).toBeCloseTo(152.10, 0);
    });

    it('lays the paragraph out by its OWN w:ind, and the hanging space stays the marker’s', async () => {
        // K states 1000/200 over a level saying 720/360, and the print puts its
        // marker at 112.10 = 50pt - 10pt: the paragraph's numbers, BOTH of
        // them. The level is not consulted at all.
        //
        // A and G are the other half of the same rule. A restates the level's
        // own indent on the paragraph and G states nothing; they print
        // identically, because a hanging indent on a list paragraph positions
        // the MARKER and never moves the text. Applying it as a first-line
        // indent as well drew every one of these items on top of its own
        // number.
        const at = await rows();

        expect(at.get('k9')!.marker).toBeCloseTo(112.10, 0);
        expect(at.get('k9')!.text).toBeCloseTo(144.10, 0);
        expect(at.get('a9')!.text).toBeCloseTo(108.10, 0);
        expect(at.get('g9')!.text).toBeCloseTo(at.get('a9')!.text, 5);
        expect(at.get('g9')!.marker).toBeCloseTo(at.get('a9')!.marker, 5);
    });

    it('measures `11.` and `10.` at the same width, as the print does', async () => {
        // A right-justified marker is placed by its own width, so this file
        // reports the width directly: LibreOffice put `10.` and `11.` at the
        // SAME 77.60, and a centred `10.` and `11.` at the same 83.85.
        //
        // This was 0.37pt out at first — exactly the GPOS kern for `1`+`1`
        // — because every string was measured kerned whether the document
        // asked or not. Kept as a DIFFERENCE rather than an absolute, because
        // half a point of tolerance is wider than that gap was and would have
        // hidden both the defect and its fix.
        const at = await rows();

        expect(at.get('b11')!.marker - at.get('b10')!.marker).toBe(0);
        expect(at.get('c11')!.marker - at.get('c10')!.marker).toBe(0);
    });
});

describe('w:sym in a font nobody here has', () => {
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const drawnBy = async () => {
        const opened = await openWordFile(file('symbol-font.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const words = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => ({ text: op.text, xPt: pt(op.xPx), yPt: pt(op.yPx) }));
        const rows = [...new Set(words.map((word) => word.yPt))].sort((a, b) => a - b);

        return rows.map((y) => words.filter((word) => y === word.yPt));
    };

    it('drops the character, room and all — where LibreOffice draws one', async () => {
        // MEASURED, NOT BUILT. Word writes bullets and dingbats as
        // `w:sym`, naming a font by name. This engine ships Liberation, Carlito
        // and Caladea and no others, so EVERY such character names a font that
        // is absent — and the reader drops it rather than print whatever
        // happens to live at that code point in a substitute, which would put
        // an `a` where the document wanted an arrow.
        //
        // LibreOffice prints something instead: the same paragraphs put `after`
        // at 104.55 for a Symbol bullet and 110.45 for a Wingdings square,
        // against 100.45 where the mark begins — so it gave them 4.10 and 10.00
        // points of room out of some substitute face.
        //
        // Here they take NO room: the symbol paragraphs lay out exactly like
        // the one with nothing between the words, which is the fourth below.
        const rows = await drawnBy();

        expect(rows[0]!.map((word) => word.text)).toEqual(['before', 'after']);
        expect(rows[1]!.map((word) => word.text)).toEqual(['before', 'after']);
        expect(rows[3]!.map((word) => word.text)).toEqual(['before', 'after']);
        expect(rows[0]![1]!.xPt).toBe(rows[3]![1]!.xPt);
    });

    it('says so, once per font it could not find', async () => {
        // The loss is reported rather than silent, which is the least a reader
        // can do about a character it will not draw — the same rule that
        // covers a text box. A literal bullet needs no such notice: U+2022 is
        // in the face already, and the third paragraph draws it at 100.35.
        const opened = await openWordFile(file('symbol-font.docx'), FONTS);
        const said = opened.document.diagnostics
            .map((entry) => entry.detail)
            .filter((detail) => detail.includes('w:sym'));

        expect(said).toEqual([
            'a w:sym in "Symbol" was dropped; the font is not available',
            'a w:sym in "Wingdings" was dropped; the font is not available',
        ]);

        const rows = await drawnBy();

        expect(rows[2]!.map((word) => word.text)).toEqual(['before', '•', 'after']);
    });
});

describe('w:noBreakHyphen, end to end', () => {
    const linesOf = async () => {
        const opened = await openWordFile(file('no-break-hyphen.docx'), FONTS);

        return layoutSections(opened.document.sections)[0]!.lines.map((line) =>
            line.line.pieces.map((piece) => piece.text).join('').trimEnd());
    };

    it('refuses the break an ordinary hyphen offers', async () => {
        // Printed by LibreOffice, the same word twice in a column
        // too narrow for it: joined by an ordinary hyphen it broke AFTER the
        // hyphen — `wwwwwwww-` then `wwwwwwww` — and joined by a
        // `w:noBreakHyphen` it did not break there at all.
        const lines = await linesOf();

        expect(lines[0]).toBe('A wwwwwwww-');
        expect(lines[1]).toBe('wwwwwwww');
        expect(lines[3]).toBe('wwwwwwww‑wwww');
    });

    it('measures the hyphen the font does not HAVE at a hyphen’s width', async () => {
        // Where the fixture earns its keep. No face this engine ships carries
        // U+2011, so it was measured as `.notdef` — 0.7778 em against a
        // hyphen's 0.3330 — and the line broke a character early: `‑www` where
        // the print has `‑wwww`.
        //
        // The third paragraph is the control: the same letters with NO hyphen
        // broke identically in both, which is what says the column width was
        // never in question and the hyphen alone was.
        const lines = await linesOf();

        expect(lines[3]).toBe('wwwwwwww‑wwww');
        expect(lines[4]).toBe('wwww');
        expect(lines[6]).toBe('wwwwwwwwwwww');
        expect(lines[7]).toBe('wwww');
    });

    it('allows a break BEFORE it, whatever Unicode says', async () => {
        // The tables forbade one, on the strength of the class Unicode gives
        // the character — and nothing had ever printed it. Measured:
        // `www ‑www`, too long for its column, broke at the SPACE and carried
        // the hyphen to the head of the next line.
        //
        // Forbidding it moved the whole fragment and then chopped it mid-word,
        // which is worse and is not what the page shows. The character refuses
        // a break after itself; that is the whole of its job.
        const lines = await linesOf();

        expect(lines[8]).toBe('D wwwwwwww');
        expect(lines[9]).toBe('‑wwwwwwww');
    });
});

describe('w:tblHeader, the row a long table repeats', () => {
    const PAGE_HEIGHT_PT = 841.89; // A4, from the printed MediaBox.

    /** Each page as the rows of its table, and where the first of them sits. */
    const pagesOf = async () => {
        const opened = await openWordFile(file('table-header-rows.docx'), FONTS);

        return layoutSections(opened.document.sections).map((page) => {
            const first = page.rows[0]!;
            const baseline = (first.yPx + first.cells[0]!.lines[0]!.baselinePx) * 72 / 96;

            return {
                topPt: Math.round((PAGE_HEIGHT_PT - baseline) * 100) / 100,
                rows: page.rows.flatMap((row) => row.cells[0]?.lines.map((line) =>
                    line.line.pieces.map((piece) => piece.text).join('').trim()) ?? []),
            };
        });
    };

    it('repeats EVERY leading header row, in order', async () => {
        // The layout has repeated headers since long before this, and no
        // fixture ever carried a `w:tblHeader` — so a feature every long table
        // in every contract leans on was built against nothing.
        //
        // Printed by LibreOffice: a 64-row table with rows one and two marked
        // opened its second page with `A01`, `A02`, then `A64`, and the first
        // of them sat at 760.94 — the top of the writing area, where the same
        // table's first page starts 750.59 because a paragraph precedes it.
        const pages = await pagesOf();

        expect(pages[1]!.rows).toEqual(['A01', 'A02', 'A64']);
        expect(pages[1]!.topPt).toBeCloseTo(760.94, 0);
        expect(pages[0]!.topPt).toBeCloseTo(750.59, 0);
    });

    it('stands above the continuation of a row it SPLIT', async () => {
        // A table breaks between its rows, so single-line rows never reach the
        // path that repeats a header above a row broken in half — which is how
        // a mutation gutting that path survived the first battery.
        //
        // A row of forty lines, started low on the page, is too tall to fit
        // whole. Printed by LibreOffice: it SPLITS — page one ends at
        // `T21` and page two opens with the repeated `HEAD` at 760.94, then
        // `T22` directly under it. So the row carries on rather than moving
        // over, and the header stands above the half of it that continues.
        const opened = await openWordFile(file('table-split-row.docx'), FONTS);
        const pages = layoutSections(opened.document.sections).map((page) => ({
            top: Math.round((PAGE_HEIGHT_PT
                - (page.rows[0]!.yPx + page.rows[0]!.cells[0]!.lines[0]!.baselinePx)
                * 72 / 96) * 100) / 100,
            rows: page.rows.flatMap((row) => row.cells[0]?.lines.map((line) =>
                line.line.pieces.map((piece) => piece.text).join('').trim()) ?? []),
        }));

        expect(pages[0]!.rows[0]).toBe('HEAD');
        expect(pages[0]!.rows[pages[0]!.rows.length - 1]).toBe('T21');
        expect(pages[1]!.rows.slice(0, 2)).toEqual(['HEAD', 'T22']);
        expect(pages[1]!.rows[pages[1]!.rows.length - 1]).toBe('T40');
        expect(pages[1]!.top).toBeCloseTo(760.94, 0);
    });

    it('splits a row AGAIN where its remainder still overflows', async () => {
        // Two defects, both found by a mutation that would not die.
        //
        // A row of 150 lines, started low: LibreOffice filled three pages —
        // `U001`..`U021`, then `U022`..`U087`, then `U088`..`U150` — repeating
        // `HEAD` above BOTH continuations. We put the whole 129-line remainder
        // on page two, running it off the bottom, and repeated no header at
        // all, because the guard weighed the header against the remainder
        // rather than against what would actually be placed.
        const opened = await openWordFile(file('table-split-row-taller.docx'), FONTS);
        const pages = layoutSections(opened.document.sections).map((page) =>
            page.rows.flatMap((row) => row.cells[0]?.lines.map((line) =>
                line.line.pieces.map((piece) => piece.text).join('').trim()) ?? []));

        expect(pages.length).toBe(3);
        expect(pages[0]!.slice(0, 2)).toEqual(['HEAD', 'U001']);
        expect(pages[0]![pages[0]!.length - 1]).toBe('U021');
        expect(pages[1]!.slice(0, 2)).toEqual(['HEAD', 'U022']);
        expect(pages[1]![pages[1]!.length - 1]).toBe('U087');
        expect(pages[2]!.slice(0, 2)).toEqual(['HEAD', 'U088']);
        expect(pages[2]![pages[2]!.length - 1]).toBe('U150');
    });

    it('moves a row that cannot be SPLIT, header and all', async () => {
        // `w:cantSplit` refuses the break. Printed by LibreOffice:
        // a 40-line row so marked put NOTHING on the page it could not fit —
        // not even the header it would have had room for — and opened the next
        // with `AHEAD` and the whole row beneath it.
        //
        // A header is never the last thing on a page, so it asks for room for
        // what it heads: the whole of an unsplittable row, one line of a
        // splittable one.
        const opened = await openWordFile(file('table-cant-split.docx'), FONTS);
        const pages = layoutSections(opened.document.sections).map((page) =>
            page.rows.flatMap((row) => row.cells[0]?.lines.map((line) =>
                line.line.pieces.map((piece) => piece.text).join('').trim()) ?? []));

        expect(pages[0]).toEqual([]);
        expect(pages[1]!.slice(0, 2)).toEqual(['AHEAD', 'A001']);
        expect(pages[1]![pages[1]!.length - 1]).toBe('A040');
    });

    it('splits one that cannot be split when NO page could hold it', async () => {
        // The flag asks for something no page can give, so it is set aside.
        // Measured in the same print: a 150-line `w:cantSplit` row spans three
        // pages, `B001`-`B066`, `B067`-`B132`, `B133`-`B150`, with `BHEAD`
        // repeated above every one.
        const opened = await openWordFile(file('table-cant-split.docx'), FONTS);
        const pages = layoutSections(opened.document.sections).map((page) =>
            page.rows.flatMap((row) => row.cells[0]?.lines.map((line) =>
                line.line.pieces.map((piece) => piece.text).join('').trim()) ?? []));

        expect(pages[2]).toEqual([]);
        expect(pages[3]!.slice(0, 2)).toEqual(['BHEAD', 'B001']);
        expect(pages[4]!.slice(0, 2)).toEqual(['BHEAD', 'B067']);
        expect(pages[5]!.slice(0, 2)).toEqual(['BHEAD', 'B133']);
        expect(pages[5]![pages[5]!.length - 1]).toBe('B150');
    });

    it('keeps two headers together, and with what they head', async () => {
        // Two header rows over a splittable row: LibreOffice kept both on the
        // page and put two lines of the row under them, then repeated both
        // above the rest.
        const opened = await openWordFile(file('table-two-headers.docx'), FONTS);
        const pages = layoutSections(opened.document.sections).map((page) =>
            page.rows.flatMap((row) => row.cells[0]?.lines.map((line) =>
                line.line.pieces.map((piece) => piece.text).join('').trim()) ?? []));

        expect(pages[1]!.slice(0, 3)).toEqual(['H1', 'H2', 'D01']);
    });

    it('repeats NONE of a header that does not lead the table', async () => {
        // The other half, and the rule the code already stated without ever
        // being asked to prove it: only a leading run of rows can be headers.
        // The same table with row TWO marked and row one not opened its second
        // page with `B64` and nothing above it.
        const pages = await pagesOf();

        expect(pages[3]!.rows).toEqual(['B64']);
        expect(pages[3]!.topPt).toBeCloseTo(760.94, 0);
    });
});

describe('wholeTable, the ninth condition', () => {
    const tablesOf = async () => {
        const opened = await openWordFile(file('whole-table.docx'), FONTS);

        return opened.document.paragraphs.filter(isTable);
    };

    it('is IGNORED, borders and shading alike — as LibreOffice ignores it', async () => {
        // MEASURED AS A NULL, which is a result rather than a gap.
        // `wholeTable` is the ninth conditional format and the one Word's own
        // built-in styles lean on: many put their borders there rather than in
        // the style's `w:tblPr`. LibreOffice draws none of it — a style whose
        // only rules are inside a `wholeTable` printed a table with no rules at
        // all, whether the rules sat in its `w:tcPr` or its `w:tblPr`, and its
        // shading never appeared either.
        //
        // So there is nothing to build: following the file would put the
        // preview out of step with the renderer that prints it, and no
        // reference here can say what the right answer would look like. Pinned
        // so the day someone reads the specification and "fixes" this, the
        // measurement answers back.
        const [cellRules, tableRules] = await tablesOf();

        expect(cellRules!.borders?.top).toBeUndefined();
        expect(cellRules!.rows[0]!.cells[0]!.borders?.top).toBeUndefined();
        expect(tableRules!.borders?.top).toBeUndefined();
    });

    it('is a real null and not a broken fixture', async () => {
        // The control, and the reason the null can be trusted: the third table
        // names a style carrying BOTH a `wholeTable` shading and a `firstRow`
        // one, over borders in the style's own `w:tblPr`. LibreOffice drew the
        // half-point rules and the first row's fill and nothing else — so the
        // machinery works on this very file, and only `wholeTable` goes unread.
        const tables = await tablesOf();
        const mixed = tables[2]!;

        expect(mixed.borders?.top?.widthPx).toBe(eighthPointsToPx(4));
        expect(mixed.rows[0]!.cells[0]!.shadingFill).toBe('#404040');
        expect(mixed.rows[1]!.cells[0]!.shadingFill).toBeUndefined();
    });
});

describe('the bands that run down the columns', () => {
    const gridsOf = async () => {
        const opened = await openWordFile(file('table-bands.docx'), FONTS);

        return opened.document.paragraphs.filter(isTable).map((table) =>
            table.rows.map((row) => row.cells.map((cell) => cell.shadingFill ?? '.')));
    };

    it('alternates the COLUMNS, and stops where noVBand says', async () => {
        // Printed by LibreOffice: `band1Vert` and `band2Vert` took
        // alternate columns down every row, and the same style under a mask
        // carrying `noVBand` took none at all.
        const [, vertical, , off] = await gridsOf();
        const stripe = ['#C0C0C0', '#909090', '#C0C0C0', '#909090'];

        expect(vertical!.every((row) => row.join() === stripe.join())).toBe(true);
        expect(off!.every((row) => row.join() === '.,.,.,.')).toBe(true);
    });

    it('beats the bands that run the other way', async () => {
        // A table defining both kinds, every cell in one of each, came out
        // ENTIRELY the vertical band's colour. That is the opposite of the
        // order ECMA-376 lists — which is why it is measured here rather than
        // read off the specification.
        const grids = await gridsOf();

        expect(grids[4]!.every((row) => row.every((fill) => '#404040' === fill))).toBe(true);
    });

    it('leaves the first column off the banding, and restarts after it', async () => {
        // The row rule mirrored, and measured rather than assumed:
        // with `firstColumn` on, the columns came out first-column, band1,
        // band2, band1 — so the column is off the count and the count begins
        // again beside it.
        const grids = await gridsOf();

        expect(grids[5]![0]).toEqual(['#404040', '#C0C0C0', '#909090', '#C0C0C0']);
        // And the same style with the column dressing only its TEXT leaves
        // column one unshaded: a band reaching under it is invisible while the
        // column carries a fill of its own, which is how a mutation removing
        // the exclusion survived its first battery — the same trap the row
        // rule hit one dimension over.
        expect(grids[6]![0]).toEqual(['.', '#C0C0C0', '#909090', '#C0C0C0']);
    });

    it('bands one row deep whatever the style asks for', async () => {
        // MEASURED, NOT BUILT. `w:tblStyleRowBandSize` says how many rows a
        // band covers; LibreOffice ignores it, banding every OTHER row for a
        // style stating two exactly as for a style stating nothing. Following
        // the file would put the preview out of step with the renderer that
        // prints it, with nothing to check the result against.
        //
        // Pinned with the one-deep answer, so the day it is built this fails.
        const [pairs] = await gridsOf();

        expect(pairs!.map((row) => row[0])).toEqual([
            '#C0C0C0', '#909090', '#C0C0C0', '#909090', '#C0C0C0', '#909090',
        ]);
    });
});

describe('w:tblStylePr, which cells each condition dresses', () => {
    /**
     * Every table in the fixture as a grid of "is this cell shaded", by the
     * fill each conditional format paints.
     *
     * The fixture gives one condition to each table and enables exactly that
     * one in the table's `w:tblLook`, so nothing overlaps and a shaded cell
     * names the condition that reached it. The seventh table is the exception,
     * and is about what happens when they DO overlap.
     */
    const gridsOf = async () => {
        const opened = await openWordFile(file('table-conditional.docx'), FONTS);

        return opened.document.paragraphs.filter(isTable).map((table) =>
            table.rows.map((row) => row.cells.map((cell) => cell.shadingFill ?? '.')));
    };

    it('gives each condition the cells LibreOffice gave it', async () => {
        // Printed, one table per condition, C0C0C0 on the cells each
        // reached: `firstRow` took row one across every column, `lastRow` the
        // last row, `firstCol` column one down every row, `lastCol` the last
        // column, and the two bands took alternate rows starting at band1.
        const [firstRow, lastRow, firstCol, lastCol, band1, band2] = await gridsOf();
        const F = '#C0C0C0';

        expect(firstRow).toEqual([[F, F, F], ['.', '.', '.'], ['.', '.', '.']]);
        expect(lastRow).toEqual([['.', '.', '.'], ['.', '.', '.'], [F, F, F]]);
        expect(firstCol).toEqual([[F, '.', '.'], [F, '.', '.'], [F, '.', '.']]);
        expect(lastCol).toEqual([['.', '.', F], ['.', '.', F], ['.', '.', F]]);
        expect(band1).toEqual([[F, F, F], ['.', '.', '.'], [F, F, F]]);
        expect(band2).toEqual([['.', '.', '.'], [F, F, F], ['.', '.', '.']]);
    });

    it('keeps a band from painting UNDER a header row', async () => {
        // A style dressing its first row's TEXT only, over a band dressing the
        // cell: LibreOffice left row one unshaded and shaded row two. So a
        // header row is off the banding altogether, rather than merely having
        // the header painted over the band — which is invisible while the
        // header carries a fill of its own, and is why this table gives it none.
        //
        // BOTH bands are defined, and that matters: with only one, a band
        // reaching under the header is invisible whenever the count puts the
        // header on the other parity — which is exactly how a mutation
        // removing this guard survived its first battery.
        const grids = await gridsOf();

        expect(grids[7]).toEqual([
            ['.', '.', '.'],
            ['#C0C0C0', '#C0C0C0', '#C0C0C0'],
            ['#909090', '#909090', '#909090'],
        ]);
    });

    it('turns everything on where there is NO w:tblLook at all', async () => {
        // Measured, not assumed: a table with no such element, over a style
        // dressing its first row and its odd bands, came out with both — the
        // header's 404040 on row one and the band's C0C0C0 on row two.
        // Reading an absent mask as nought, which is the tempting answer,
        // would have dressed neither.
        const grids = await gridsOf();

        expect(grids[8]).toEqual([
            ['#404040', '#404040', '#404040'],
            ['#C0C0C0', '#C0C0C0', '#C0C0C0'],
            ['.', '.', '.'],
        ]);
    });

    it('carries BORDERS as well as shading, under the cell’s own', async () => {
        // The commonest table look there is: a heavy rule under the header.
        // Printed by LibreOffice, a style lending its first row a 3pt
        // bottom over a table drawing half a point everywhere — the rule under
        // row one came out 3.000 and every other 0.500.
        //
        // And a cell saying `none` still wins: the same style over such a cell
        // drew the table's own 0.500 under row one, because the cell drew no
        // bottom at all and the row below supplied its top.
        // Asserted on the rules DRAWN rather than on the cell's own borders:
        // a cell that refuses its bottom still has one drawn under it, by the
        // table, and the model would report only the refusal.
        const opened = await openWordFile(file('table-conditional-borders.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        // Widest per HEIGHT, because both cells of a shared edge draw it —
        // the row above its bottom and the row below its top, at the same y.
        // Taking them in order finds the thin one as readily as the thick.
        const widest = new Map<number, number>();
        for (const op of renderPage(page).ops) {
            if ('line' !== op.kind || op.y1Px !== op.y2Px) {
                continue;
            }
            const y = Math.round(op.y1Px * 100) / 100;
            widest.set(y, Math.max(widest.get(y) ?? 0, op.widthPx));
        }
        const heights = [...widest.entries()].sort((left, right) => left[0] - right[0]);
        // Four rules to a table: its top, two between its rows, its bottom.
        const underHeader = (index: number): number =>
            Math.round(heights[index * 4 + 1]![1] * 72 / 96 * 100) / 100;

        expect(underHeader(0)).toBeCloseTo(3, 1);
        expect(underHeader(1)).toBeCloseTo(0.5, 1);
        expect(underHeader(2)).toBeCloseTo(0.5, 1);
    });

    it('settles which one wins where they MEET', async () => {
        // A seventh table with a header, a first column and a band together,
        // each shaded differently. Row one came out the header's 404040 in
        // EVERY column — so a row condition beats a column one — and row two
        // came out the column's 808080 in the first cell and the band's
        // C0C0C0 in the rest, so a column condition beats a band.
        //
        // And the bands count the BODY: row two is band1 here, where the table
        // above with no header made row ONE band1.
        const grids = await gridsOf();

        expect(grids[6]).toEqual([
            ['#404040', '#404040', '#404040'],
            ['#808080', '#C0C0C0', '#C0C0C0'],
            ['#808080', '.', '.'],
        ]);
    });
});

describe('what else a table style carries, end to end', () => {
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const wordsOf = async () => {
        const opened = await openWordFile(file('table-style-text.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const words = renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => ({ text: op.text.trim(), xPt: pt(op.xPx), yPt: pt(op.yPx) }));

        // Exact matching where a tag is a prefix of the cell tags around it:
        // the rows are drawn before the body lines, so a loose 'C' finds a
        // cell of the table rather than the paragraph naming it.
        return (tag: string, exact = false) => words.find(
            (word) => exact ? word.text === tag : word.text.startsWith(tag))!;
    };

    it('dresses the text in the style’s own w:rPr', async () => {
        // Printed by LibreOffice: a style whose `w:rPr` says fourteen
        // points stepped its rows 16.60, where the same table naming no style
        // stepped 12.05 — a 14pt line against a 10pt one, plus the half point
        // of rule between them. The cells themselves say nothing about size.
        //
        // It changes the geometry, not just the look: a table whose text is
        // four points bigger than we drew it paginates somewhere else.
        const word = await wordsOf();

        expect(Math.round((word('A2').yPt - word('A1').yPt) * 100) / 100).toBeCloseTo(16.60, 0);
        expect(Math.round((word('E2').yPt - word('E1').yPt) * 100) / 100).toBeCloseTo(12.05, 0);
    });

    it('lays the paragraphs out by the style’s own w:pPr', async () => {
        // The same, for the paragraph half, printed by LibreOffice:
        // a style saying centred put its text at 183.05, where the tables
        // naming no such style put theirs at 77.25.
        const word = await wordsOf();

        expect(word('B1').xPt).toBeCloseTo(183.05, 0);
        expect(word('E1').xPt).toBeCloseTo(77.25, 0);
    });

    it('lets a style BASED ON another climb down from it', async () => {
        // A leaf restating the size at eleven points over a root saying
        // fourteen stepped its rows 13.15 — the eleven — so the chain answers
        // leaf-first here exactly as it does for a table style's borders.
        const word = await wordsOf();

        expect(Math.round((word('F2').yPt - word('F1').yPt) * 100) / 100).toBeCloseTo(13.15, 0);
    });

    it('leaves the dress INSIDE the table it belongs to', async () => {
        // The paragraph after every table, printed at 72.10 — left-aligned and
        // plain, where the centred table above would have put it at 183.05 if
        // what a table lends outlived it. A table nested in another cell is the
        // same question one level down, which is why the lend is saved and put
        // back rather than simply set.
        const word = await wordsOf();

        // The paragraph after the CENTRED table, which is the one a leak would
        // move: matched on its first word alone, because the display list draws
        // a line word by word and no op holds the whole of 'C header on'.
        expect(word('C', true).xPt).toBeCloseTo(72.10, 0);
        expect(word('G').xPt).toBeCloseTo(72.10, 0);
    });

    it('dresses the first row where the LOOK enables it, and not where it does not', async () => {
        // Measured and pinned unbuilt earlier; this is the build.
        //
        // Printed by LibreOffice: the header table's first row came out 16.55
        // TALL — the fill drawn behind it — for a 14pt bold line where the
        // body is 10pt, and its baselines sat 12.85 apart. The height and the
        // step are different quantities, and asserting the first against the
        // second is how this test failed at 12.87 while the engine was right.
        //
        // The same table with the mask cleared steps the body's own, because
        // the gate is the hex mask and not the attributes beside it.
        const word = await wordsOf();
        const step = (tag: string): number => Math.round(
            (word(tag + '2').yPt - word(tag + '1').yPt) * 100) / 100;

        expect(step('C')).toBeCloseTo(12.85, 0);
        // Close rather than equal: these baselines are rounded before they are
        // subtracted, so one spacing prints as 11.99 in one table and 12.00 in
        // another (an earlier table comparison made the same mistake in the
        // same way).
        expect(step('D')).toBeCloseTo(step('E'), 1);
    });

    it('shades the first row from the conditional format, and only it', async () => {
        // The fill LibreOffice drew behind that row, C0C0C0, 233.95 wide and
        // 16.55 tall — and nothing behind the same table with the mask cleared.
        const opened = await openWordFile(file('table-style-text.docx'), FONTS);
        const tables = opened.document.paragraphs.filter(isTable);
        const fillOf = (index: number, row: number): string | undefined =>
            tables[index]!.rows[row]!.cells[0]!.shadingFill;

        expect(fillOf(2, 0)).toBe('#C0C0C0');
        expect(fillOf(2, 1)).toBeUndefined();
        expect(fillOf(3, 0)).toBeUndefined();
    });
});

describe('a paragraph split by a page break, end to end', () => {
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    /** The words on each page, and where the first of them sits. */
    const pagesOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections).map((page) => ({
            topPt: pt(page.lines[0]!.yPx + page.lines[0]!.baselinePx),
            words: page.lines.map((line) =>
                line.line.pieces.map((piece) => piece.text).join('').trim()),
        }));
    };

    it('gives an EMPTY piece of the split no line at all', async () => {
        // Printed by LibreOffice, four arrangements of one break:
        // alone in its paragraph, after text, before text, and between two
        // runs of text. Every page after a break opened at the same 760.49 —
        // 81.40 from the top — so an empty piece takes no room on either side.
        //
        // The first arrangement is how PHPWord writes `addPageBreak()`, so it
        // is in every document this platform generates: each opened with a
        // blank line at the top of every page after the first.
        const pages = await pagesOf('page-break-empty.docx');

        expect(pages.map((page) => page.words)).toEqual([
            ['A top'],
            ['A next', 'B top', 'B tail'],
            ['B next', 'C top'],
            ['C head', 'C next', 'D top', 'D tail'],
            ['D head', 'D next'],
        ]);
        // Every page starting at the SAME height is the whole of the finding:
        // a blank line anywhere would show as one page starting 11.50 lower.
        for (const page of pages) {
            expect(page.topPt).toBeCloseTo(81.40, 0);
        }
    });

    it('still breaks the page when the piece it drew nothing for carried it', async () => {
        // The trap in fixing this. `pageBreakBefore` rides on every piece after
        // the first, so a piece that is dropped takes the break with it and
        // runs the pages together. Kept with no RUNS instead: the flow gives it
        // no line and still honours the break.
        expect((await pagesOf('page-break-empty.docx')).length).toBe(5);
    });
});

describe('a REAL document this platform generated, against its print', () => {
    /**
     * A lease agreement from the document-generation pipeline: PHPWord, whose
     * `docProps/app.xml` says so outright.
     *
     * Small, and it earns its place by being nothing like the other real
     * fixture — LANDSCAPE, Cyrillic headings, bookmarks, and its pages
     * separated by a break-alone paragraph. It matched to a hundredth of a
     * point once that was read the way LibreOffice reads it,
     * so it is held to a tenth rather than the point and a half the longer
     * document needs.
     */
    const PAGE_HEIGHT_PT = 595.30; // A4 landscape, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    // Baselines LibreOffice printed, measured DOWN from the top of the page.
    const PRINTED = [
        [81.40, 92.90],
        [81.40, 92.90],
        [81.40],
    ];

    it('places every line of every page within a tenth of a point', async () => {
        const opened = await openWordFile(file('real-lease.docx'), FONTS);
        const pages = layoutSections(opened.document.sections);

        expect(pt(pages[0]!.geometry.heightPx)).toBeCloseTo(PAGE_HEIGHT_PT, 1);
        expect(pages.length).toBe(PRINTED.length);

        const deviations = pages.flatMap((page, index) => page.lines.map((line, row) =>
            Math.abs(pt(line.yPx + line.baselinePx) - PRINTED[index]![row]!)));

        expect(deviations.length).toBe(5);
        expect(Math.max(...deviations)).toBeLessThan(0.1);
    });
});

describe('a REAL document, against the page LibreOffice printed', () => {
    /**
     * Not a probe of one rule: a whole document, printed and compared.
     *
     * 160 paragraphs and four tables, laid out here and put beside the
     * baselines LibreOffice printed from the same file — every line of them,
     * vendored beside the fixture. Four fixes came out of this one comparison
     * — the default family, the room a table keeps for its rules, the default
     * size, a bare heading's spacing — each uncovered only once the one before
     * it stopped masking it.
     *
     * It guards what a probe cannot: that the rules hold TOGETHER, on a file
     * nobody built to test them.
     */
    const PAGE_HEIGHT_PT = 792; // US Letter, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    const printed = JSON.parse(readFileSync(
        join(FIXTURES, 'real-adr.libreoffice.json'), 'utf8')) as {
            pageHeightPt: number;
            pages: number[][];
            /** Where each of those lines BEGINS — see the x test below. */
            left: number[][];
        };

    const ours = async (): Promise<number[][]> => {
        const opened = await openWordFile(file('real-adr.docx'), FONTS);

        return layoutSections(opened.document.sections).map((page) => {
            const baselines = new Set<number>();
            for (const op of renderPage(page).ops) {
                if ('text' === op.kind) {
                    baselines.add(Math.round(pt(op.yPx) * 100) / 100);
                }
            }

            return [...baselines].sort((left, right) => left - right);
        });
    };

    /**
     * The leftmost x drawn on each line, in the same order as {@link ours}.
     *
     * EMPTY text is skipped and BLANK text is not, which is not a detail: a
     * code line here begins with four spaces, and this engine draws them as a
     * piece of their own where the print keeps them inside the line's one
     * operation. Skipping ours reported the line 21.50pt to the right — four
     * monospace spaces at 9pt — and it was the instrument that was wrong.
     */
    const oursLeft = async (inkOnly = false): Promise<number[][]> => {
        const opened = await openWordFile(file('real-adr.docx'), FONTS);

        return layoutSections(opened.document.sections).map((page) => {
            const rows = new Map<number, { x: number; ink: boolean }>();
            for (const op of renderPage(page).ops) {
                if ('text' !== op.kind || '' === op.text) {
                    continue;
                }

                const y = Math.round(pt(op.yPx) * 100) / 100;
                const seen = rows.get(y);
                rows.set(y, {
                    x: Math.min(seen?.x ?? Infinity, pt(op.xPx)),
                    ink: (seen?.ink ?? false) || '' !== op.text.trim(),
                });
            }

            return [...rows.entries()]
                .sort((left, right) => left[0] - right[0])
                .filter(([, row]) => row.ink || !inkOnly)
                .map(([, row]) => row.x);
        });
    };

    /**
     * How many lines begin in each column, to the nearest half point.
     *
     * Half a point because that is the tolerance the row-by-row comparison
     * uses, and because the print runs a constant 0.05 to 0.10 right of this
     * engine — a whole-page offset, not a layout difference, and one that must
     * not split a column in two.
     */
    const census = (values: readonly number[]): Map<number, number> => {
        const counts = new Map<number, number>();
        for (const value of values) {
            const column = Math.round(value * 2) / 2;
            counts.set(column, (counts.get(column) ?? 0) + 1);
        }

        return counts;
    };

    it('paginates it in THREE pages, as LibreOffice did', async () => {
        expect(printed.pageHeightPt).toBe(PAGE_HEIGHT_PT);
        expect((await ours()).length).toBe(printed.pages.length);
    });

    it('places every line of page one within a point and a half of the print', async () => {
        // The drift is cumulative and one-directional — 0.08 at the title,
        // 1.34 by the last line, about 0.045 a line. That is LibreOffice's own
        // line rounding: it steps 11.55 for a 10pt Liberation Serif line where
        // the font's metrics give 11.4990, and 10.20 for a 9pt one where they
        // give 10.195. We keep the font's number, as everywhere else a renderer
        // rounds what a file states.
        //
        // Held to a point and a half so the guard is about the RULES holding
        // together. Every fault this comparison found was worth ten of it: the
        // font alone was 12pt by the second line.
        const pages = await ours();
        const deviations = pages[0]!
            .slice(0, printed.pages[0]!.length)
            .map((value, index) => Math.abs(value - printed.pages[0]![index]!));

        expect(Math.max(...deviations)).toBeLessThan(1.5);
    });

    it('starts every line of pages one and two where the print starts it', async () => {
        // The other axis, and the one this comparison was blind to for its
        // whole life. Baselines cannot see a horizontal defect at
        // all — not a tab in the wrong column, not an indent applied
        // twice. Two slices in a row moved x on every document in the
        // corpus, and this file, the strictest guard the engine has, could not
        // have failed either time.
        //
        // Half a point, not the point and a half the baselines are given: there
        // is no cumulative drift ACROSS a line, so each one either starts in
        // the right column or does not.
        //
        // Pages one and two only, and ordinally, which works for one reason:
        // each of them holds exactly one line more than the print — the
        // difference pinned below — and on both it is the LAST. Page three
        // therefore starts two lines further into the document than the print's
        // does and cannot be lined up row by row; the census below is what
        // covers it, and its four table columns.
        const pages = await oursLeft();

        for (const index of [0, 1]) {
            const want = printed.left[index]!;
            const deviations = pages[index]!
                .slice(0, want.length)
                .map((value, row) => Math.abs(value - want[row]!));

            expect(deviations.length).toBe(want.length);
            expect(Math.max(...deviations)).toBeLessThan(0.5);
        }
    });

    it('uses the same COLUMNS as the print, over the whole document', async () => {
        // What the row-by-row comparison above cannot reach: page three, where
        // all four tables are. Counting how many lines begin in each column
        // does not care which page they landed on, so the page-break difference
        // stops mattering and the tables are covered.
        //
        // Every column matches exactly — the body margin, the two list indents,
        // and the five columns the tables put text in — except the margin
        // itself, where this engine draws one line more. That is the same
        // difference as below: a line the print holds over, drawn here.
        const ours = census((await oursLeft(true)).flat());
        const theirs = census(printed.left.flat());
        const columns = [...new Set([...ours.keys(), ...theirs.keys()])].sort((a, b) => a - b);

        expect(columns).toEqual([72, 77.5, 90, 194.5, 233.5, 311.5, 389.5, 428.5]);
        for (const column of columns) {
            expect([column, ours.get(column)])
                .toEqual([column, (theirs.get(column) ?? 0) + (72 === column ? 1 : 0)]);
        }
    });

    it('fits ONE line more on page one than the print does', async () => {
        // Pinned rather than fixed, and it is the drift above rather than a
        // rule of its own: by the foot of the page LibreOffice is 1.34 lower,
        // which is enough to push its last code line over. Ours keeps it.
        //
        // So this is what the rounding costs — a page break in a different
        // place — and the day the line height matches, this test fails and the
        // number becomes zero.
        const pages = await ours();

        expect(pages[0]!.length - printed.pages[0]!.length).toBe(1);
    });
});

describe('a built-in heading that states no spacing, end to end', () => {
    const pt = (px: number): number => px * 72 / 96;

    /**
     * The gap above and below the paragraph whose text starts with `text`.
     *
     * Baseline to baseline on either side, which is what the probe reads —
     * and the heading's own size cancels out of the comparison, since every
     * pair below is between two paragraphs of the SAME size.
     */
    const aroundIn = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const pages = layoutSections(opened.document.sections);
        const lines = pages.flatMap((page, index) => page.lines.map((line) => ({
            text: line.line.pieces.map((piece) => piece.text).join('').trim(),
            // Stacked page after page, so a gap that straddles a break is not
            // read as a negative number.
            at: index * pages[0]!.geometry.heightPx + line.yPx + line.baselinePx,
        })));

        return (text: string): { before: number; after: number } => {
            const index = lines.findIndex((line) => line.text.startsWith(text));
            const round = (value: number): number => Math.round(pt(value) * 100) / 100;

            return {
                before: round(lines[index]!.at - lines[index - 1]!.at),
                after: round(lines[index + 1]!.at - lines[index]!.at),
            };
        };
    };

    it('spaces it 12pt above and 6pt below, where nothing else says', async () => {
        // Measured against LibreOffice with a control beside it: two
        // stub heading styles of the SAME size, one silent about `w:spacing`
        // and one stating zero on both sides.
        //
        //   silent   24.50 above, 17.70 below
        //   zero     12.50 above, 11.70 below
        //
        // The difference is the whole of it — 12.00 and 6.00 — and it is what
        // put every heading in an under-specified document 12pt high and every
        // paragraph after one 6pt high.
        const around = await aroundIn('heading-defaults.docx');
        // Within the half point this file allows everywhere: the residual is
        // LibreOffice's own line rounding, 11.55 to our 11.50 at ten points.
        // What is exact is the DIFFERENCE between the two, which is the rule.
        expect(around('H4 head').before).toBeCloseTo(24.50, 0);
        expect(around('H4 head').after).toBeCloseTo(17.70, 0);
        expect(around('H5 head').before).toBeCloseTo(12.50, 0);
        expect(around('H5 head').after).toBeCloseTo(11.70, 0);
        expect(Math.round((around('H4 head').before - around('H5 head').before) * 100) / 100)
            .toBe(12);
        expect(Math.round((around('H4 head').after - around('H5 head').after) * 100) / 100)
            .toBe(6);
    });

    it('knows a heading by its NAME, not by its style id', async () => {
        // Measured both ways round, which is the only way to tell them apart:
        // a style called `Zebra` and named `heading 6` was spaced like a
        // heading, and one called `Heading7` and named `Custom Thing` was not.
        const around = await aroundIn('heading-defaults.docx');

        expect(around('ZZ head')).toEqual(around('H4 head'));
        expect(around('CC head')).toEqual(around('H5 head'));
        expect(around('ZZ head').before).toBeCloseTo(24.50, 0);
        expect(around('CC head').before).toBeCloseTo(12.50, 0);
    });

    it('stands aside where the document supplies spacing of its own', async () => {
        // The half that makes this safe. The same file with a real `Normal`
        // stating 20pt above and 10pt below: LibreOffice printed the style
        // named `heading 8`, the one named `Custom Thing` and the one named
        // `heading 6` at EXACTLY the same distances — 32.45 and 31.75 — so the
        // built-in does not fire when anything else answers.
        //
        // An ordinary Word document, whose Normal always carries spacing, never
        // reaches the fallback at all.
        const around = await aroundIn('heading-defaults-inherited.docx');

        expect(around('II head')).toEqual(around('CC head'));
        expect(around('II head')).toEqual(around('ZZ head'));
        expect(around('II head').before).toBeCloseTo(32.45, 0);
        expect(around('II head').after).toBeCloseTo(31.75, 0);
    });
});

describe('the room a table keeps for its rules, end to end', () => {
    const pt = (px: number): number => px * 72 / 96;

    /**
     * The pitch from one row's text to the next, per table on the page.
     *
     * Read off BASELINES rather than off the row boxes, because that is what
     * the probe reads: a row can be the right height and still be put in the
     * wrong place, which is exactly the fault this fixture is about.
     */
    const pitchesOf = async () => {
        const opened = await openWordFile(file('table-rule-room.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const byTable = new Map<number, number[]>();

        for (const row of page.rows) {
            const baseline = row.cells[0]?.lines[0]?.baselinePx;
            if (undefined === baseline) {
                continue;
            }
            const key = row.blockIndex;
            // Kept at full precision and rounded only where the PITCH is
            // taken: rounding each baseline first makes one spacing print as
            // two different numbers depending where on the page it falls.
            byTable.set(key, [...(byTable.get(key) ?? []), row.yPx + baseline]);
        }

        return [...byTable.values()].map((baselines) => baselines
            .slice(1)
            .map((value, index) => Math.round(pt(value - baselines[index]!) * 100) / 100));
    };

    it('keeps room for the rule the CELL declares, not the one the table does', async () => {
        // Printed by LibreOffice, five tables of three rows, each row
        // one 9pt line — a 10.35 box — so the pitch IS the line plus the rule:
        //
        //   A  table 1/2pt, cells 1/8pt   10.45   the cell's, thinner
        //   B  table 1/2pt, cells silent  10.85   the table's, nothing overrides
        //   C  table 1/8pt, cells 3pt     13.35   the cell's, thicker
        //   D  cells 1/8pt and 3pt        13.35   the widest of the two
        //   E  3pt under row 1, 1/8 over row 2 -> 13.35, then 10.45
        //
        // A is the shape of the real document that gave the fault away: we
        // drew an eighth of a point and reserved half of one, every row, all
        // the way down.
        //
        // The eighth-point rows are 10.47 here against that 10.45, and the
        // 0.025 is LibreOffice quantising a hairline: its own stroke widths
        // report 0.100 for a rule the file declares as `w:sz="1"`, an eighth of
        // a point. We keep the eighth the document states, as everywhere else a
        // renderer rounds what a file says. Every case where the RULE differs —
        // half a point, three points, the widest of two, the widest across an
        // edge — lands exactly.
        expect(await pitchesOf()).toEqual([
            [10.47, 10.47],
            [10.85, 10.85],
            [13.35, 13.35],
            [13.35, 13.35],
            [13.35, 10.47],
        ]);
    });

    it('keeps the same room BELOW the last row, where the text starts again', async () => {
        // The outer edges follow the cell too, which the stroke widths show
        // outright: table A declares half a point and its printed rules — top
        // and bottom included — came out at the cells' hairline. So the
        // paragraph after a table starts a cell's rule below it, not a table's.
        //
        // Measured from the last row's baseline to the next paragraph's:
        // 10.45 after table A, whose cells state an eighth, and 10.85 after
        // table B, whose cells state nothing at all.
        const opened = await openWordFile(file('table-rule-room.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const gapAfter = (blockIndex: number): number => {
            const rows = page.rows.filter((row) => blockIndex === row.blockIndex);
            const last = rows[rows.length - 1]!;
            const bottom = last.yPx + last.cells[0]!.lines[0]!.baselinePx;
            const below = page.lines
                .map((line) => line.yPx + line.baselinePx)
                .filter((value) => value > bottom)
                .sort((left, right) => left - right)[0]!;

            return Math.round(pt(below - bottom) * 100) / 100;
        };

        expect(gapAfter(1)).toBe(10.47);
        expect(gapAfter(3)).toBe(10.85);
    });
});

describe('the font a document that names none gets, end to end', () => {
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const linesOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return { opened, page };
    };

    it('takes Times New Roman where a styles part states no family', async () => {
        // Measured with the FONT'S OWN NAME, read out of the printed PDF rather
        // than inferred from a step: a `.docx` whose `styles.xml`
        // carries an empty `w:rPrDefault` embedded `LiberationSerif`, and
        // nothing else. Its 20pt paragraphs stepped 23.00 — the 1.15 em of that
        // face, where Carlito's 1.2207 would have given 24.41.
        //
        // Not a corner case: an empty `docDefaults` is what an uploaded
        // template that states nothing looks like, and the fill path copies its
        // `styles.xml` verbatim — so answering Calibri here made such a
        // document paginate one way in the preview and another in the renderer
        // that prints it.
        const { page } = await linesOf('no-font-stated.docx');
        const twenties = page.lines.filter((line) => line.line.pieces.some(
            (piece) => piece.text.startsWith('twenty')));

        expect(twenties).toHaveLength(2);
        expect(pt(twenties[1]!.yPx - twenties[0]!.yPx)).toBeCloseTo(23.00, 0);
        expect(twenties[0]!.line.pieces[0]!.font.familyName).toBe('Liberation Serif');
    });

    it('keeps Calibri where there is no styles part AT ALL', async () => {
        // The other half, and the reason this is a rule about the styles part
        // rather than about fonts in general: the same probe with no
        // `styles.xml` in the package embedded `Carlito-Regular`. LibreOffice
        // treats a missing styles part and an empty one differently, and so
        // does this — the manifest's own default answers the first.
        const { opened } = await linesOf('cell-text-turned-wrapped.docx');
        const table = opened.document.paragraphs.filter(isTable)[0]!;
        const cell = table.rows[0]!.cells[0]!;
        const paragraph = cell.paragraphs[0]!;

        expect(isTable(paragraph) ? '' : paragraph.runs[0]!.font.familyName).toBe('Carlito');
    });

    it('sets a run that states no SIZE at ten points, beside a ten-point ruler', async () => {
        // The other half of the same fallback. Paragraphs with no
        // `w:rPr` at all, and 10pt paragraphs beside them in the same print:
        // LibreOffice stepped 11.55 for both, so a run that says nothing is
        // set at ten points. We answered eleven, which is a point of drift on
        // every blank line a document leaves between its blocks.
        const { page } = await linesOf('no-size-stated.docx');
        const step = (first: number, second: number): number =>
            Math.round(pt(page.lines[second]!.yPx - page.lines[first]!.yPx) * 100) / 100;

        expect(step(0, 1)).toBe(step(3, 4));
        expect(step(0, 1)).toBeCloseTo(11.55, 0);
    });

    it('keeps ELEVEN points where there is no styles part at all', async () => {
        // The control that decides where the fallback belongs, printed by
        // LibreOffice from the same probe as a bare package: it
        // stepped 13.45 between its unstated paragraphs and 12.20 between its
        // 10pt ones — eleven points and ten points of Carlito's 1.2207 em. So
        // the manifest's `defaults` are right as they stand, and the ten-point
        // answer belongs beside the Times New Roman one rather than in place
        // of them.
        const { page } = await linesOf('no-size-stated-bare.docx');
        const step = (first: number, second: number): number =>
            Math.round(pt(page.lines[second]!.yPx - page.lines[first]!.yPx) * 100) / 100;

        expect(step(0, 1)).toBeCloseTo(13.45, 0);
        expect(step(3, 4)).toBeCloseTo(12.20, 0);
    });

    it('says nothing about a substitution it made up itself', async () => {
        // A document that never named a font did not ask for Times New Roman —
        // this engine did. Routing the fallback through the substitution path
        // would put a "not available, using the metric-compatible…" notice on
        // every under-specified document, and a diagnostic that fires on
        // everything is worse than none.
        const { opened } = await linesOf('no-font-stated.docx');

        expect(opened.document.diagnostics).toEqual([]);
    });
});

describe('w:gutter, end to end', () => {
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    /**
     * One entry per page: where its text starts, and what its lines hold.
     *
     * The fixture's first four sections are one page each, and the fifth — the
     * two-column one — runs on past them, so the tests below take the four
     * they are about rather than counting pages.
     */
    const pagesOf = async () => {
        const opened = await openWordFile(file('page-gutter.docx'), FONTS);

        return layoutSections(opened.document.sections).map((page) => ({
            leftPt: pt(page.lines[0]!.xPx),
            headerLeftPt: pt(page.header!.lines[0]!.xPx),
            footerLeftPt: pt(page.footer!.lines[0]!.xPx),
            // The SECOND line, which is the wrapping paragraph's first: how
            // many words fit says whether the writing width shrank, and a text
            // that merely slid right would keep all of them.
            words: page.lines[1]!.line.pieces
                .map((piece) => piece.text).join('').trim().split(/\s+/).length,
        }));
    };

    it('adds the binding margin to the LEFT margin, and takes it off the width', async () => {
        // Printed by LibreOffice, four sections of one page each: no
        // gutter, half an inch, half an inch in an RTL section, and a whole
        // inch. The text began at 72.10, 108.10, 107.90 and 144.10 — the shift
        // following the value — and the same paragraph broke at a different
        // word each time: its section tag plus 22 words of `mm` on the first
        // page, then 20, 20 and 18. Counted WITH the tag here, which is what
        // the printed line holds.
        //
        // Both halves matter. A gutter that moved the text without narrowing
        // the column would keep every page's line at 23 and print the last
        // word past the right margin.
        const pages = await pagesOf();

        expect(pages.slice(0, 4).map((page) => page.leftPt)).toEqual([72, 108, 108, 144]);
        expect(pages.slice(0, 4).map((page) => page.words)).toEqual([23, 21, 21, 19]);
    });

    it('moves the header and footer with it', async () => {
        // What makes the gutter part of the page's margin rather than an
        // indent on the body: every section drew its furniture at exactly the
        // body's own left edge, all four times.
        const pages = await pagesOf();

        for (const page of pages.slice(0, 4)) {
            expect(page.headerLeftPt).toBe(page.leftPt);
            expect(page.footerLeftPt).toBe(page.leftPt);
        }
    });

    it('keeps the gutter on the LEFT of an RTL section, as LibreOffice does', async () => {
        // Word puts the binding edge on the right for a `w:bidi` section.
        // LibreOffice does not — it printed the third page at 107.90, within a
        // fifth of a point of the second — and neither do we. Pinned rather
        // than argued: this engine does not lay out right-to-left text at all,
        // so a gutter that changed sides would be the only bidi-aware thing in
        // it, and would move the text of a document it cannot otherwise read.
        const pages = await pagesOf();

        expect(pages[2]!.leftPt).toBe(pages[1]!.leftPt);
    });

    it('lays a two-column section inside the gutter, not across it', async () => {
        // The last section is two columns with half an inch of binding margin.
        // LibreOffice put the first at 108.10 and the second at 333.75: the
        // writing area less the gutter, split in two about a 36pt space. A
        // gutter that moved the text without narrowing the area would have put
        // the second column half an inch further right and hung it off the page.
        const opened = await openWordFile(file('page-gutter.docx'), FONTS);
        const columned = layoutSections(opened.document.sections)[4]!;
        const edges = [...new Set(columned.lines.map((line) => pt(line.xPx)))]
            .sort((left, right) => left - right);

        expect(edges.length).toBe(2);
        expect(edges[0]!).toBeCloseTo(108.10, 0);
        expect(edges[1]!).toBeCloseTo(333.75, 0);
    });
});

describe('content controls around a table, end to end', () => {
    const PAGE_HEIGHT_PT = 841.89; // A4, from the printed MediaBox.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    const opened = async () => openWordFile(file('sdt-table.docx'), FONTS);

    /** Every drawn word, with where it landed, in PDF coordinates. */
    const drawn = async () => {
        const page = layoutSections((await opened()).document.sections)[0]!;

        return renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => ({
                text: op.text.trim(),
                xPt: pt(op.xPx),
                yPt: Math.round((PAGE_HEIGHT_PT - pt(op.yPx)) * 100) / 100,
            }));
    };

    it('keeps the cells of a row inside a control, where LibreOffice keeps them', async () => {
        // A repeating section is how a template says "one row per line item",
        // so this is the shape a generated document actually has. Printed by
        // LibreOffice: the wrapped row's `BB1` and `BB2` at 77.25 and
        // 311.25, in the middle row of three, exactly like the plain row above.
        const words = await drawn();
        const at = (text: string) => words.find((word) => text === word.text)!;

        expect(at('BB1').xPt).toBeCloseTo(77.25, 0);
        expect(at('BB2').xPt).toBeCloseTo(311.25, 0);
        expect(at('BB1').yPt).toBeCloseTo(736.49, 0);
        // The rows either side of it, to prove the wrapped one took its place
        // in the table rather than merely being drawn somewhere.
        expect(at('AA1').yPt).toBeCloseTo(748.49, 0);
        expect(at('CC1').yPt).toBeCloseTo(724.49, 0);
    });

    it('draws a cell inside a control, which LIBREOFFICE leaves blank', async () => {
        // A DELIBERATE divergence. LibreOffice keeps the cell — its inside
        // vertical rule runs the full height of all three rows — and prints
        // nothing in it, so `CC2` is missing from its page. The document says
        // the cell holds that text, and losing a cell's words is the fault
        // already closed for the text box; it is not reopened here.
        //
        // Dropping the cell outright, which is what this reader used to do, is
        // worse than either: the row then has ONE cell where the table has two
        // columns, so everything after a control shifts a column left.
        const words = await drawn();
        const second = words.find((word) => 'CC2' === word.text)!;

        expect(second.xPt).toBeCloseTo(311.25, 0);
        expect(second.yPt).toBeCloseTo(724.49, 0);
    });

    it('measures a wrapped table exactly as it measures the same table unwrapped', async () => {
        // The property that made this a defect rather than a preference: two
        // walks over one table have to agree. `readTable` flattened controls
        // and `columnWidths` did not, so a table whose only row was wrapped
        // kept the NOMINAL grid its file declares — 100 twips, six pixels to a
        // column — while the identical table without the wrapper recovered
        // 4680 twips from its cells.
        //
        // The two tables here differ by the wrapper and by nothing else, which
        // is what makes this answerable without asking LibreOffice: it renders
        // BOTH at the nominal grid, because it trusts a grid where this engine
        // recomputes it from the cells — a divergence that predates this slice.
        const tables = (await opened()).document.paragraphs.filter(isTable);
        const [, wrapped, unwrapped] = tables;

        expect(tables.length).toBe(4);
        expect(wrapped!.columnWidthsPx).toEqual(unwrapped!.columnWidthsPx);
        expect(wrapped!.columnWidthsPx).toEqual([twipsToPx(4680), twipsToPx(4680)]);

        const words = await drawn();
        const xOf = (text: string) => words.find((word) => text === word.text)!.xPt;

        expect([xOf('DD1'), xOf('DD2')]).toEqual([xOf('EE1'), xOf('EE2')]);
    });

    it('counts a controlled cell as a COLUMN, not only as content', async () => {
        // Skipping a cell costs twice, and the second cost is the quiet one:
        // the column counter never advances, so every cell after the control
        // is credited to the column before its own. Here the second cell is
        // the only thing that states the second column's width, and the two
        // are deliberately far apart — 7020 and 2340 twips — so a cell
        // credited to the wrong column cannot look like a rounding difference.
        const table = (await opened()).document.paragraphs.filter(isTable)[3]!;

        expect(table.columnWidthsPx).toEqual([twipsToPx(7020), twipsToPx(2340)]);
    });
});

describe('w:lvlOverride, end to end', () => {
    /** Every list marker on the page, in the order they are drawn. */
    const markersOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);

        return layoutSections(opened.document.sections)[0]!.lines
            .map((line) => line.marker?.run.text)
            .filter((marker): marker is string => undefined !== marker);
    };

    it('counts three numIds on ONE abstract straight through', async () => {
        // Printed by LibreOffice: 1. 2. 3. for the plain list, 7. 8.
        // 9. for the one told to start at seven, and (J) (K) (L) for the one
        // whose override replaces the level — the tenth, eleventh and twelfth
        // letters, because the count carried on from the 9 above it.
        //
        // We used to print 1. 2. 3. three times over: a counter per numId, and
        // both overrides unread.
        expect(await markersOf('list-override.docx')).toEqual([
            '1.', '2.', '3.',
            '7.', '8.', '9.',
            '(J)', '(K)', '(L)',
        ]);
    });

    it('follows the same three lists INTERLEAVED', async () => {
        // The same definitions in a different order, which is what separates
        // the rules from each other. `(C)` proves a replacement level's own
        // `w:start` is the initial value; `5.` proves the count is the
        // abstract's; `11.` proves the override is spent after its first use.
        expect(await markersOf('list-override-order.docx')).toEqual([
            '(C)', '(D)',
            '5.', '6.',
            '7.', '8.',
            '9.', '10.',
            '11.', '12.',
        ]);
    });
});

describe('text boxes, end to end', () => {
    const PAGE_HEIGHT_PT = 841.89; // A4, from the printed MediaBox.
    const pt = (px: number): number => Math.round(px * 72 / 96 * 100) / 100;

    /**
     * Every line the page draws, as `{ x, y, text }` in PDF coordinates.
     *
     * Gathered by BASELINE because a line is drawn word by word, and turned up
     * the other way because that is how the probe reads them: comparing our
     * numbers with LibreOffice's means putting both in the same frame.
     */
    const linesOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;
        const byLine = new Map<number, { xPt: number; text: string }>();

        for (const op of renderPage(page).ops) {
            if ('text' !== op.kind) {
                continue;
            }
            const seen = byLine.get(op.yPx);
            byLine.set(op.yPx, {
                xPt: seen?.xPt ?? pt(op.xPx),
                text: (seen?.text ?? '') + op.text,
            });
        }

        return [...byLine.entries()].map(([yPx, line]) => ({
            xPt: line.xPt,
            yPt: Math.round((PAGE_HEIGHT_PT - pt(yPx)) * 100) / 100,
            text: line.text.trim(),
        }));
    };

    it('draws a DrawingML box where LibreOffice draws it, and moves nothing else', async () => {
        // An earlier pass measured this file and reported the loss rather than
        // fixing it. This is the fix: the words inside `w:txbxContent` are
        // read as blocks, stacked at the frame's width and drawn at the frame's
        // corner — so what was silently missing is now on the page.
        //
        // LibreOffice printed `DMLBOX` at 279.30, 702.39. The frame is 200pt
        // across the column and 20pt down the paragraph, and the 279.30 is the
        // column's own left plus that 200, plus the 7.20 default inset; the
        // 0.10 we are under it is the glyph side bearing the probe reads and
        // the display list does not — the same 0.1pt every measured line in
        // this file allows for, which is what `toBeCloseTo(x, 0)` is here
        // for. A box placed against the wrong origin would be out by 200.
        const lines = await linesOf('text-box.docx');
        const box = lines.find((line) => line.text.startsWith('DMLBOX'))!;

        expect(box.xPt).toBeCloseTo(279.30, 0);
        expect(box.yPt).toBeCloseTo(702.39, 0);
        // And the body is undisturbed: five paragraphs stepping their ordinary
        // 11.50, exactly as LibreOffice printed them. A box that pushed the
        // text about would be a worse fault than one that lost it, because the
        // damage would be to words the author can see.
        const body = lines.filter((line) => !line.text.includes('BOX'));

        expect(body.map((line) => line.text))
            .toEqual(['before', 'vml anchor', 'middle', 'dml anchor', 'after']);
        expect(body.map((line) => Math.round((body[0]!.yPt - line.yPt) * 100) / 100))
            .toEqual([0, 11.50, 23, 34.50, 46]);
    });

    it('breaks the text at the box WIDTH, less the inset the file states', async () => {
        // The measurement that decided the model: the same fourteen
        // words in two 120pt boxes, one silent about its inset and one stating
        // zero on all four sides. LibreOffice broke the silent one five words
        // to a line and the zero one six, which is the 7.20 either side of it —
        // so a box's width is what breaks its text, and the inset comes off it
        // before anything is measured.
        const lines = await linesOf('text-box-inside.docx');
        const inset = lines.filter((line) => line.text.startsWith('DA') || 'mm' === line.text.slice(0, 2)
            && line.xPt > 279);
        const flush = lines.filter((line) => line.text.startsWith('DZ')
            || 'mm' === line.text.slice(0, 2) && line.xPt < 273);

        expect(inset.slice(0, 3).map((line) => line.text)).toEqual([
            'DA mm mm mm mm',
            'mm mm mm mm mm',
            'mm mm mm mm mm',
        ]);
        expect(flush.slice(0, 3).map((line) => line.text)).toEqual([
            'DZ mm mm mm mm mm',
            'mm mm mm mm mm mm',
            'mm mm mm',
        ]);

        // Both corners where LibreOffice put them: 279.30/736.89 for the box
        // that takes the default inset, and 272.10/628.99 for the one that
        // waives it — the frame's own edge, 200pt from the column.
        expect(inset[0]!.xPt).toBeCloseTo(279.30, 0);
        expect(inset[0]!.yPt).toBeCloseTo(736.89, 0);
        expect(flush[0]!.xPt).toBeCloseTo(272.10, 0);
        expect(flush[0]!.yPt).toBeCloseTo(628.99, 0);
        // And the lines inside step the body's own 11.50: a box needs no
        // leading rule of its own, which is the other half of what was measured.
        expect(Math.round((inset[0]!.yPt - inset[1]!.yPt) * 100) / 100).toBe(11.50);
    });

    it('believes a VML inset that LIBREOFFICE ignores, and lands 2.85 off it', async () => {
        // A DELIBERATE divergence, and the one place this file disagrees with
        // the renderer it is measured against.
        //
        // LibreOffice printed both VML boxes at 276.35 — the one silent about
        // its inset AND the one stating `inset="0,0,0,0"`. Stating zero changed
        // nothing, so that 4.25 is LibreOffice's own default rather than
        // anything the document asked for. We read what the file says: the
        // format's 7.20 when it is silent, and zero when it says zero.
        //
        // The evidence that this is right is in LibreOffice's own print of the
        // DrawingML pair, where `lIns="0"` DID move the text to the frame's
        // edge at 272.10 — the same edge our zero-inset VML box lands on.
        const lines = await linesOf('text-box-inside.docx');
        const silent = lines.find((line) => line.text.startsWith('VA'))!;
        const zero = lines.find((line) => line.text.startsWith('VZ'))!;

        // Ours, against LibreOffice's 276.35 for both. The gap is exactly the
        // inset difference on each axis and nothing else, which is what makes
        // it a disagreement about one number rather than a placement bug.
        expect(silent.xPt).toBeCloseTo(279.20, 0);
        expect(zero.xPt).toBeCloseTo(272, 0);
        expect(Math.round((silent.xPt - zero.xPt) * 100) / 100).toBe(7.20);
    });

    it('leaves a PICTURE a picture, and says nothing about a box', async () => {
        // The other half of the check: without it, treating every drawing as a
        // text box passes everything above and quietly turns each picture in
        // each document into an empty float.
        const opened = await openWordFile(file('inline-image.docx'), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        expect(renderPage(page).ops.some((op) => 'image' === op.kind)).toBe(true);
        expect(opened.document.diagnostics
            .filter((entry) => entry.detail.includes('text box'))).toEqual([]);
    });
});

describe('w:shd behind text, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;

    const rectsOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return renderPage(page).ops
            .filter((op): op is RectOp => 'rect' === op.kind)
            .map((op) => ({
                fill: op.fill.toLowerCase(),
                xPx: pt(op.xPx),
                footPx: PAGE_HEIGHT_PT - pt(op.yPx + op.heightPx),
                widthPx: pt(op.widthPx),
                heightPx: pt(op.heightPx),
            }));
    };

    it('paints a RUN’s shading as the line’s height by the run’s own advance', async () => {
        // `w:rPr/w:shd` is the other way a document says "colour behind this
        // text", beside `w:highlight`, and the reader has folded it onto the
        // same field all along — untested against a printed page until now.
        //
        // Printed: `re 85.900 712.439 41.650 11.450` for a yellow run
        // and `re 84.800 689.439 27.650 11.450` for a red one — the LINE's
        // height by the RUN's advance, starting 0.10 before the glyphs, which
        // is the same shape measured for a highlight in a turned cell.
        const rects = await rectsOf('paragraph-run-shading.docx');
        const yellow = rects.find((rect) => '#ffff00' === rect.fill);
        const red = rects.find((rect) => '#ff0000' === rect.fill);

        expect(Math.abs(yellow!.xPx - 85.90)).toBeLessThan(0.5);
        expect(Math.abs(yellow!.footPx - 712.439)).toBeLessThan(0.5);
        expect(Math.abs(yellow!.widthPx - 41.65)).toBeLessThan(0.5);
        expect(Math.abs(yellow!.heightPx - 11.45)).toBeLessThan(0.5);
        expect(Math.abs(red!.widthPx - 27.65)).toBeLessThan(0.5);
    });

    it('paints NO shading for a PARAGRAPH’s w:shd, which is what LibreOffice does', async () => {
        // Not a gap — parity, and a null result with its control in the same
        // conversion, which is the rule for a null. The fixture's first
        // paragraph asks for grey and its third for green, both with `w:shd`
        // in schema order before `w:ind`; LibreOffice paints neither. What it
        // DOES paint is the
        // run shading in the same document, which is the control: shading
        // reaches the converter, the drawing XML is sound, and only the
        // paragraph's is declined.
        //
        // Word paints one. This engine follows the reference it can measure,
        // and there is nothing to measure a box against while the reference
        // draws none — not its width, not where it starts.
        const rects = await rectsOf('paragraph-run-shading.docx');

        expect(rects.map((rect) => rect.fill).sort()).toEqual(['#ff0000', '#ffff00']);
    });
});

describe('w:ruby, end to end', () => {
    const PAGE_HEIGHT_PT = 841.861; // A4, from the printed MediaBox.
    const pt = (px: number): number => px * 72 / 96;
    const upFromFoot = (px: number): number => PAGE_HEIGHT_PT - pt(px);

    /** Every drawn word of the page, with where and how big it was drawn. */
    const drawnOf = async (name: string) => {
        const opened = await openWordFile(file(name), FONTS);
        const page = layoutSections(opened.document.sections)[0]!;

        return renderPage(page).ops
            .filter((op): op is TextOp => 'text' === op.kind && '' !== op.text.trim())
            .map((op) => ({
                text: op.text.trim(),
                xPx: pt(op.xPx),
                yPx: upFromFoot(op.yPx),
                sizePx: pt(op.sizePx),
            }));
    };

    it('draws the gloss above the base, on a line grown to hold it', async () => {
        // The gloss was READ and dropped for a while, on the reasoning that
        // it "needs a second line inside one line's height" — which is exactly
        // what LibreOffice gives it. Printed: a 5pt gloss over a 10pt base put
        // the base's baseline at 708.74 where the paragraph above it sits at
        // 725.99, so the line is 17.25 rather than 11.50 — the base's own line
        // plus the GLOSS's 5.75, all of it above the baseline.
        //
        // The gloss then sits at the top of that room on a baseline of its own,
        // 4.68 below the line's top, which is its own ascent at 5pt.
        const drawn = await drawnOf('ruby-gloss.docx');
        const gloss = drawn.find((op) => 'gloss' === op.text);
        const base = drawn.find((op) => 'base' === op.text);

        expect(gloss).toBeDefined();
        expect(Math.abs(gloss!.yPx - 719.19)).toBeLessThan(0.5);
        expect(Math.abs(gloss!.xPx - 91.00)).toBeLessThan(0.5);
        expect(Math.abs(gloss!.sizePx - 5)).toBeLessThan(0.1);
        expect(Math.abs(base!.yPx - 708.74)).toBeLessThan(0.5);
    });

    it('leaves the lines after it plain, and the paragraphs after it where they were', async () => {
        // The growth belongs to the glossed LINE: the second line of
        // the same paragraph printed 697.24, a plain 11.50 below the first, and
        // the paragraph after it at 685.74.
        const drawn = await drawnOf('ruby-gloss.docx');
        const second = drawn.filter((op) => 'hh' === op.text)[1];
        const after = drawn.filter((op) => 'gap2' === op.text)[0];

        expect(Math.abs(second!.yPx - 697.24)).toBeLessThan(0.5);
        expect(Math.abs(after!.yPx - 685.74)).toBeLessThan(0.5);
    });

    it('lets a gloss WIDER than its base decide the advance, and centres the base', async () => {
        // Printed: a gloss of eight `w` at 5pt over an `xy` base — the gloss
        // from x=86.00, the base at 95.45 centred under it, and the text after
        // the group at 114.85, which is the gloss's own width along rather than
        // the base's. A group measured by its base alone would have drawn the
        // gloss over whatever followed.
        const drawn = await drawnOf('ruby-gloss.docx');
        const gloss = drawn.find((op) => op.text.startsWith('www'));
        const base = drawn.find((op) => 'xy' === op.text);
        // The run after the group begins with a space, so the printed 114.85 is
        // the space and this is the first letter past it. Picked by the LINE it
        // is on: `cc` appears in three paragraphs of this fixture, and counting
        // them found the wrong one.
        const following = drawn.find((op) => 'cc' === op.text && Math.abs(op.yPx - 668.49) < 0.5);

        expect(Math.abs(gloss!.xPx - 86.00)).toBeLessThan(0.5);
        expect(Math.abs(base!.xPx - 95.45)).toBeLessThan(0.5);
        expect(Math.abs(following!.xPx - 117.35)).toBeLessThan(0.5);
    });
});
