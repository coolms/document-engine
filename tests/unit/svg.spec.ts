import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import { XmlDocument } from '../../src/ooxml/xml.js';
import {
    DEFAULT_PARAGRAPH_STYLE,
    layoutPages,
    type PageGeometry,
    type Paragraph,
} from '../../src/layout/page-layout.js';
import { renderPageToSvg, serialise } from '../../src/render/svg.js';
import type { TextOp } from '../../src/render/display-list.js';
import { layoutSections } from '../../src/layout/page-layout.js';
import { FontCatalogue, type FontManifest } from '../../src/word/font-catalogue.js';
import { openWordFile } from '../../src/word/word-package.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MONO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Regular.ttf'))));
const BOLD = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Bold.ttf'))));
const ITALIC = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Italic.ttf'))));

const PAGE: PageGeometry = {
    widthPx: 400,
    heightPx: 200,
    marginTopPx: 10,
    marginRightPx: 10,
    marginBottomPx: 10,
    marginLeftPx: 10,
};

function para(text: string, font = MONO): Paragraph {
    return {
        runs: [{ text, font, sizePx: 16 }],
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: 20 },
    };
}

const svgOf = (paragraph: Paragraph, options = {}): string =>
    renderPageToSvg(layoutPages([paragraph], PAGE)[0]!, options);

/**
 * All the line's text, joined.
 *
 * A line is one `<text>` per PIECE, and a piece is a word with the spaces that
 * follow it — so "a  b" is two elements, not one.
 */
const drawnText = (svg: string): string =>
    XmlDocument.parse(svg).root.elements('text').map((element) => element.text).join('');

describe('renderPageToSvg', () => {
    it('produces XML this package can parse back', () => {
        // The engine has an XML parser, so the SVG is checked by READING it
        // rather than by matching a string — a serialiser tested against its
        // own expected output only proves it did not change.
        const document = XmlDocument.parse(svgOf(para('hello')));

        expect(document.root.name).toBe('svg');
        expect(document.root.attribute('viewBox')).toBe('0 0 400 200');
        expect(document.root.elements('text').length).toBe(1);
    });

    it('names the face the engine MEASURED with', () => {
        const text = XmlDocument.parse(svgOf(para('x'))).root.element('text');

        expect(text?.attribute('font-family')).toBe('Liberation Mono');
        expect(text?.attribute('font-size')).toBe('16');
    });

    it('reads weight and slant off the FILE, since they are separate files here', () => {
        // A run cannot say it is bold — boldness is which font was chosen. A
        // renderer that ignored this would draw a bold heading in the regular
        // face while the layout measured the bold one.
        const attributes = (font: TrueTypeFont) => {
            const text = XmlDocument.parse(svgOf(para('x', font))).root.element('text');

            return [text?.attribute('font-weight'), text?.attribute('font-style')];
        };

        expect(attributes(MONO)).toEqual([null, null]);
        expect(attributes(BOLD)).toEqual(['bold', null]);
        expect(attributes(ITALIC)).toEqual([null, 'italic']);
    });

    it('keeps spaces that would otherwise be collapsed', () => {
        // A piece can start or end with a space, and losing it shifts every
        // word after it left of where it was measured.
        const svg = svgOf(para('a  b'));

        expect(XmlDocument.parse(svg).root.element('text')?.attribute('xml:space'))
            .toBe('preserve');
        expect(drawnText(svg)).toBe('a  b');
    });

    it('escapes text that would otherwise break the document', () => {
        const svg = svgOf(para('<a & b>'));

        expect(svg).not.toContain('<a & b>');
        expect(drawnText(svg)).toBe('<a & b>');
    });

    it('draws paper under the text, and none when asked for none', () => {
        expect(XmlDocument.parse(svgOf(para('x'))).root.element('rect')?.attribute('fill'))
            .toBe('#ffffff');
        expect(XmlDocument.parse(svgOf(para('x'), { backgroundColor: null })).root.element('rect'))
            .toBeNull();
    });

    it('carries supplied CSS through so a page can be made self-contained', () => {
        // `@font-face` rules are the point: without them the SVG asks for a
        // family the viewer may not have.
        const css = '@font-face{font-family:"Liberation Mono";src:url(x.ttf)}';
        const svg = svgOf(para('x'), { styleCss: css });

        expect(svg).toContain(`<![CDATA[${css}]]>`);
        expect(XmlDocument.parse(svg).root.element('style')).not.toBeNull();
    });

    it('serialises shading as a rect and a border as a stroke', () => {
        const side = { widthPx: 2, colorHex: '#123456', style: 'solid' as const };
        const table = {
            rows: [{
                cells: [{ paragraphs: [para('x')], gridSpan: 1, shadingFill: '#EEEEEE' }],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [200],
            borders: { top: side },
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };
        const root = XmlDocument.parse(
            renderPageToSvg(layoutPages([table], PAGE)[0]!),
        ).root;

        // The first rect is the paper; the second is the shading.
        expect(root.elements('rect')[1]?.attribute('fill')).toBe('#EEEEEE');
        expect(root.element('line')?.attribute('stroke')).toBe('#123456');
        expect(root.element('line')?.attribute('stroke-width')).toBe('2');
        expect(root.element('line')?.attribute('stroke-dasharray')).toBeNull();
    });

    it('dashes and dots a border that asked for it', () => {
        const dashes = (style: 'dashed' | 'dotted'): string | null => {
            const table = {
                rows: [{
                    cells: [{ paragraphs: [para('x')], gridSpan: 1 }],
                    isHeader: false,
                    cantSplit: false,
                }],
                columnWidthsPx: [200],
                borders: { top: { widthPx: 2, colorHex: '#000000', style } },
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            };

            return XmlDocument.parse(renderPageToSvg(layoutPages([table], PAGE)[0]!))
                .root.element('line')?.attribute('stroke-dasharray') ?? null;
        };

        // Both scale with the stroke, so a heavy dashed border keeps its rhythm.
        expect(dashes('dashed')).toBe('6,4');
        expect(dashes('dotted')).toBe('2,4');
    });

    it('draws a run in its own colour, and the rest in the default', () => {
        const coloured: Paragraph = {
            runs: [
                { text: 'aa', font: MONO, sizePx: 16, colorHex: '#FF0000' },
                { text: 'bb', font: MONO, sizePx: 16 },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: 20 },
        };
        const fills = (options = {}): (string | null)[] =>
            XmlDocument.parse(renderPageToSvg(layoutPages([coloured], PAGE)[0]!, options))
                .root.elements('text').map((element) => element.attribute('fill'));

        expect(fills()).toEqual(['#FF0000', '#000000']);
        // The default follows the page; the run's own colour does not.
        expect(fills({ textColor: '#FFFFFF' })).toEqual(['#FF0000', '#FFFFFF']);
    });
});

describe('a run turned on its side', () => {
    // `w:textDirection`. The layout that decides WHICH runs turn is a second
    // measuring pass and is not built; this is the renderer's half, so
    // that when the layout lands there is nowhere for it to be dropped — which
    // is precisely what happened to the line numbers, the page border and the
    // cell's own box before anyone went looking.
    const turned = (turn: 'ccw' | 'cw' | undefined, xPx: number, yPx: number): TextOp => ({
        kind: 'text',
        xPx,
        yPx,
        text: 'a',
        font: MONO,
        sizePx: 16,
        ...(undefined === turn ? {} : { turn }),
    });

    it('turns anticlockwise for `ccw` and clockwise for `cw`', () => {
        // Measured off the PDF's text matrix: LibreOffice writes `btLr` as
        // `[0 1 -1 0]` and `tbRl` as `[0 -1 1 0]` — a quarter turn each way
        // about the run's own baseline origin, which is what `rotate(±90 x y)`
        // does in SVG.
        const svg = serialise({
            widthPx: 100,
            heightPx: 100,
            pageNumber: 1,
            ops: [turned('ccw', 40, 60), turned('cw', 70, 80), turned(undefined, 10, 20)],
        });

        expect(svg).toContain('transform="rotate(-90 40 60)"');
        expect(svg).toContain('transform="rotate(90 70 80)"');
    });

    it('leaves an upright run without a transform at all', () => {
        const svg = serialise({
            widthPx: 100, heightPx: 100, pageNumber: 1, ops: [turned(undefined, 10, 20)],
        });

        expect(svg).not.toContain('transform');
    });
});

describe('what reaches the SVG', () => {
    // Every slice of this arc verified the engine's own model and its draw
    // operations. None checked the LAST hop — the SVG the preview and the
    // artifact are actually made of. A feature read, laid out, and dropped by
    // the serialiser looks exactly like one that was never built: the same
    // shape as the walker that collected nothing and the placement
    // that discarded what it collected.
    const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), '../fixtures/docx');
    const MANIFEST = JSON.parse(
        readFileSync(join(FONT_DIR, 'fonts.manifest.json'), 'utf8')) as FontManifest;
    const FONTS = FontCatalogue.load(
        MANIFEST, (name) => new Uint8Array(readFileSync(join(FONT_DIR, name))));

    const svgOf = async (name: string, index = 0): Promise<string> => {
        const opened = await openWordFile(
            readFileSync(join(FIXTURES, name)), FONTS);

        return renderPageToSvg(layoutSections(opened.document.sections)[index]!);
    };

    it('draws the numbers down the margin', async () => {
        // `w:lnNumType` becomes text operations like any other, so a
        // serialiser that only knew rules and rectangles would lose them.
        const svg = await svgOf('line-numbers.docx');
        const digits = [...svg.matchAll(/<text[^>]*>(\d+)<\/text>/gu)].map((m) => m[1]);

        expect(digits).toEqual(['1', '2', '3', '4', '5']);
    });

    it('draws the page’s own border', async () => {
        const svg = await svgOf('page-border-from-page.docx');

        // Four rules and no others: the page border is the only thing this
        // fixture strokes.
        expect([...svg.matchAll(/<line /gu)]).toHaveLength(4);
    });

    it('draws a bordered paragraph inside a table cell', async () => {
        const svg = await svgOf('cell-paragraph-border.docx');
        const strokes = [...svg.matchAll(/<line /gu)];

        // The table's own four sides and the cell's box: more than a table
        // alone would draw.
        expect(strokes.length).toBeGreaterThan(4);
        expect(svg).toContain('BOXED');
    });

    it('draws a run’s own box, and the text inside it', async () => {
        const svg = await svgOf('run-border.docx');

        // Four boxed runs, four rules apiece.
        expect([...svg.matchAll(/<line /gu)]).toHaveLength(16);
        expect(svg).toContain('MID');
    });

    it('draws a table’s rules once each, as the display list emits them', async () => {
        // The count the renderer settled on survives serialisation: one rule
        // per row plus a top, not two per shared edge.
        const svg = await svgOf('table-rule-gaps.docx');
        const horizontal = [...svg.matchAll(/<line x1="([\d.]+)" y1="([\d.]+)" x2="([\d.]+)" y2="([\d.]+)"/gu)]
            .filter((match) => match[2] === match[4]);

        // By POSITION: a two-column table draws each rule as two collinear
        // segments, one per cell across its own width. LibreOffice emits one
        // segment for the pair; abutting halves paint the same ink, and the
        // claim that matters is that there are four rules and not eight.
        const rules = [...new Set(horizontal.map((match) => match[2]))];

        expect(rules).toHaveLength(4);
        expect(horizontal).toHaveLength(8);
    });
});
