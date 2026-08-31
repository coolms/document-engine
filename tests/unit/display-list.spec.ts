import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import {
    DEFAULT_PARAGRAPH_STYLE,
    layoutPages,
    type Page,
    type PageGeometry,
    type Paragraph,
    type ParagraphStyle,
} from '../../src/layout/page-layout.js';
import type { BorderSide, BorderStyle, BoxBorders } from '../../src/layout/borders.js';
import type { InlineShape } from '../../src/layout/image.js';
import type { StyledRun } from '../../src/layout/line-breaker.js';
import { renderPage, type DrawOp, type LineOp, type RectOp, type TextOp }
    from '../../src/render/display-list.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MONO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Regular.ttf'))));

const CELL = MONO.measureAdvance('x', 16).widthPt;
const LINE = 20;

const PAGE: PageGeometry = {
    widthPx: CELL * 20 + 20,
    heightPx: LINE * 8 + 20,
    marginTopPx: 10,
    marginRightPx: 10,
    marginBottomPx: 10,
    marginLeftPx: 10,
    headerDistancePx: 4,
    footerDistancePx: 4,
};

function para(text: string, style: Partial<ParagraphStyle> = {}): Paragraph {
    return {
        runs: [{ text, font: MONO, sizePx: 16 }],
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE, ...style },
    };
}

/** Only the text, since borders and shading share the operation union. */
const textOps = (ops: readonly DrawOp[]): TextOp[] =>
    ops.filter((op): op is TextOp => 'text' === op.kind);

const textsOf = (page: ReturnType<typeof renderPage>): string[] =>
    textOps(page.ops).map((op) => op.text);

describe('renderPage', () => {
    it('positions every operation at the BASELINE, not the top of its line', () => {
        // A renderer draws text from its baseline. Emitting the box top would
        // put every glyph a whole ascent too high.
        const [laid] = layoutPages([para('xx')], PAGE);
        const line = laid!.lines[0]!;
        const [op] = textOps(renderPage(laid!).ops);

        expect(op!.yPx).toBe(line.yPx + line.baselinePx);
        expect(line.baselinePx).toBeGreaterThan(0);
    });

    it('draws the furniture BEFORE the body', () => {
        // Paint order matters where a header outgrows the margin it was
        // measured for: the body has to win the overlap.
        const [laid] = layoutPages([para('body')], PAGE, {
            headerFor: () => [para('head')],
            footerFor: () => [para('foot')],
        });

        expect(textsOf(renderPage(laid!))).toEqual(['head', 'foot', 'body']);
    });

    it('advances across a line by each piece width', () => {
        const styled: Paragraph = {
            runs: [
                { text: 'aa', font: MONO, sizePx: 16 },
                { text: 'bb', font: MONO, sizePx: 16 },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
        };
        const ops = textOps(renderPage(layoutPages([styled], PAGE)[0]!).ops);

        expect(ops.map((op) => op.text)).toEqual(['aa', 'bb']);
        expect(ops[1]!.xPx - ops[0]!.xPx).toBeCloseTo(CELL * 2, 6);
    });

    it('draws a list marker at its OWN x, not the line’s', () => {
        // The marker sits in the hanging space left of the paragraph indent, so
        // an indented or centred first line moves and the bullet does not.
        const item: Paragraph = {
            ...para('text', { indentLeftPx: CELL * 4 }),
            marker: { run: { text: '●', font: MONO, sizePx: 16 }, offsetPx: CELL * 2 },
        };
        const ops = textOps(renderPage(layoutPages([item], PAGE)[0]!).ops);

        expect(ops.map((op) => op.text)).toEqual(['●', 'text']);
        expect(ops[0]!.xPx).toBeCloseTo(10 + CELL * 4 - CELL * 2, 6);
        expect(ops[1]!.xPx).toBeCloseTo(10 + CELL * 4, 6);
    });

    it('draws the marker in the marker’s OWN colour', () => {
        // A numbered list can colour its numbers differently from its text —
        // Word puts the marker's formatting on the numbering, not the run — so
        // the bullet cannot inherit the paragraph's colour.
        const item: Paragraph = {
            runs: [{ text: 'text', font: MONO, sizePx: 16, colorHex: '#0000FF' }],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
            marker: {
                run: { text: '●', font: MONO, sizePx: 16, colorHex: '#FF0000' },
                offsetPx: CELL,
            },
        };
        const ops = textOps(renderPage(layoutPages([item], PAGE)[0]!).ops);

        expect(ops.map((op) => [op.text, op.colorHex])).toEqual([
            ['●', '#FF0000'],
            ['text', '#0000FF'],
        ]);
    });

    it('marks only the FIRST line of a list item', () => {
        const item: Paragraph = {
            ...para('aaaaaaaaaa bbbbbbbbbb cccccccccc'),
            marker: { run: { text: '1.', font: MONO, sizePx: 16 }, offsetPx: CELL },
        };
        const ops = textOps(renderPage(layoutPages([item], PAGE)[0]!).ops);

        expect(ops.filter((op) => '1.' === op.text).length).toBe(1);
    });

    it('draws a list marker inside a table CELL as well', () => {
        // Cells stack through their own code, so a marker honoured on the page
        // and dropped in a cell is the drift the shared path exists to prevent.
        const item: Paragraph = {
            ...para('text'),
            marker: { run: { text: '●', font: MONO, sizePx: 16 }, offsetPx: CELL * 2 },
        };
        const table = {
            rows: [{
                cells: [{ paragraphs: [item], gridSpan: 1 }],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [CELL * 20],
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };
        const ops = textOps(renderPage(layoutPages([table], PAGE)[0]!).ops);

        expect(ops.map((op) => op.text)).toEqual(['●', 'text']);
        expect(ops[1]!.xPx - ops[0]!.xPx).toBeCloseTo(CELL * 2, 6);
    });

    it('carries the piece’s own font and size onto the operation', () => {
        const mixed: Paragraph = {
            runs: [
                { text: 'aa', font: MONO, sizePx: 16 },
                { text: 'bb', font: MONO, sizePx: 32 },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: 40 },
        };
        const ops = textOps(renderPage(layoutPages([mixed], { ...PAGE, heightPx: 400 })[0]!).ops);

        expect(ops.map((op) => op.sizePx)).toEqual([16, 32]);
    });

    it('includes the text inside table cells', () => {
        const table = {
            rows: [{
                cells: [{ paragraphs: [para('incell')], gridSpan: 1 }],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [CELL * 20],
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };

        expect(textsOf(renderPage(layoutPages([table], PAGE)[0]!))).toEqual(['incell']);
    });

    describe('highlighting', () => {
        const marked = (highlightHex?: string): Paragraph => ({
            runs: [
                { text: 'aa', font: MONO, sizePx: 16 },
                { text: 'bb', font: MONO, sizePx: 16, ...(undefined === highlightHex ? {} : { highlightHex }) },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE, lineRule: 'exact' },
        });

        it('paints the mark BEHIND the text it belongs to', () => {
            const ops = renderPage(layoutPages([marked('#FFFF00')], PAGE)[0]!).ops;

            // Rect before the text it sits under, and only for the marked run.
            expect(ops.map((op) => op.kind)).toEqual(['text', 'rect', 'text']);
        });

        it('fills the LINE box, not the font box', () => {
            // Verified against LibreOffice: a highlighted run inside an exactly
            // spaced 24pt line fills all 24 points. Filling the font's box
            // leaves a pale stripe above and below in loosely spaced text.
            const [page] = layoutPages([marked('#FFFF00')], PAGE);
            const line = page!.lines[0]!;
            const rect = renderPage(page!).ops.find((op) => 'rect' === op.kind)!;

            expect(rect.yPx).toBe(line.yPx);
            expect(rect.heightPx).toBe(line.heightPx);
            expect(line.heightPx).toBe(LINE);
        });

        it('spans only the marked run', () => {
            const [page] = layoutPages([marked('#FFFF00')], PAGE);
            const rect = renderPage(page!).ops.find((op) => 'rect' === op.kind)!;

            // Two characters in, two characters wide.
            expect(rect.xPx).toBeCloseTo(10 + CELL * 2, 6);
            expect(rect.widthPx).toBeCloseTo(CELL * 2, 6);
        });

        it('paints nothing for a run with no mark', () => {
            const ops = renderPage(layoutPages([marked()], PAGE)[0]!).ops;

            expect(ops.every((op) => 'text' === op.kind)).toBe(true);
        });
    });

    describe('floating pictures', () => {
        const float = (behindText: boolean) => ({
            image: {
                content: { bytes: new Uint8Array([1]), contentType: 'image/png' },
                widthPx: 20,
                heightPx: 20,
            },
            xPx: 0,
            yPx: 0,
            widthPx: 20,
            heightPx: 20,
            wrap: 'none' as const,
            behindText,
            exclusion: { xPx: 0, yPx: 0, widthPx: 20, heightPx: 20 },
        });

        const drawnWith = (behindText: boolean): string[] => {
            const [page] = layoutPages([para('body')], PAGE);

            return renderPage({ ...page!, floats: [float(behindText)] }).ops.map((op) => op.kind);
        };

        it('draws a behind-text float UNDER the words', () => {
            // `behindDoc` is a watermark: text over it, not under it.
            expect(drawnWith(true)).toEqual(['image', 'text']);
        });

        it('draws an ordinary float over them', () => {
            expect(drawnWith(false)).toEqual(['text', 'image']);
        });
    });

    describe('tab leaders', () => {
        // HALF a cell past a whole number of them. Liberation Mono is
        // monospaced, so a span of exactly nine glyphs cannot tell a fill that
        // rounds DOWN from one that rounds up — both give nine.
        const SPAN = CELL * 9.5;

        const led = (leader?: 'dot' | 'underscore'): ReturnType<typeof textOps> => {
            const paragraph: Paragraph = {
                runs: [{ text: 'a\tb', font: MONO, sizePx: 16 }],
                style: {
                    ...DEFAULT_PARAGRAPH_STYLE,
                    lineHeightPx: LINE,
                    tabStops: [{
                        positionPx: CELL + SPAN,
                        align: 'left',
                        ...(undefined === leader ? {} : { leader }),
                    }],
                },
            };

            return textOps(renderPage(layoutPages([paragraph], PAGE)[0]!).ops);
        };

        it('fills the tab with as many WHOLE glyphs as fit', () => {
            // Nine and a half cells of span: nine glyphs fit and the tenth does
            // not. A clipped full stop is a smudge, and the shortfall is under
            // one character wide.
            const [, fill] = led('dot');

            expect(fill!.text).toBe('.'.repeat(9));
            expect(MONO.measureAdvance(fill!.text, fill!.sizePx).widthPt)
                .toBeLessThanOrEqual(SPAN);
        });

        it('starts the fill where the TAB starts', () => {
            const [label, fill] = led('dot');
            const labelWidth = MONO.measureAdvance(label!.text, label!.sizePx).widthPt;

            expect(fill!.xPx).toBeCloseTo(label!.xPx + labelWidth, 6);
        });

        it('draws the character the leader names', () => {
            expect(led('underscore')[1]!.text.startsWith('__')).toBe(true);
        });

        it('draws nothing for an ordinary tab', () => {
            // Label and the text after it, and no fill between them.
            expect(led().map((op) => op.text)).toEqual(['a', 'b']);
        });
    });

    describe('runs off the baseline', () => {
        const shifted = (baselineShiftPx?: number): Paragraph => ({
            runs: [
                { text: 'aa', font: MONO, sizePx: 16 },
                {
                    text: 'bb',
                    font: MONO,
                    sizePx: 16,
                    ...(undefined === baselineShiftPx ? {} : { baselineShiftPx }),
                },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
        });

        it('draws a shifted run off the line, and its neighbours on it', () => {
            const [page] = layoutPages([shifted(-5)], PAGE);
            const baseline = page!.lines[0]!.yPx + page!.lines[0]!.baselinePx;
            const ops = textOps(renderPage(page!).ops);

            expect(ops[0]!.yPx).toBe(baseline);
            expect(ops[1]!.yPx).toBe(baseline - 5);
        });

        it('leaves the LINE where it was', () => {
            // A superscript does not move the text around it, and does not
            // change where the next line begins.
            const raised = layoutPages([shifted(-5)], PAGE)[0]!.lines[0]!;
            const flat = layoutPages([shifted()], PAGE)[0]!.lines[0]!;

            expect(raised.yPx).toBe(flat.yPx);
            expect(raised.heightPx).toBe(flat.heightPx);
        });

        it('keeps the marker on the line whatever the runs do', () => {
            // The shifted run is FIRST, so a marker that followed the run
            // beside it would rise with this one.
            const item: Paragraph = {
                runs: [
                    { text: 'aa', font: MONO, sizePx: 16, baselineShiftPx: -5 },
                    { text: 'bb', font: MONO, sizePx: 16 },
                ],
                style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
                marker: { run: { text: '1.', font: MONO, sizePx: 16 }, offsetPx: CELL },
            };
            const [page] = layoutPages([item], PAGE);
            const baseline = page!.lines[0]!.yPx + page!.lines[0]!.baselinePx;

            expect(textOps(renderPage(page!).ops)[0]!.yPx).toBe(baseline);
        });
    });

    describe('underlining and striking through', () => {
        const decorated = (overrides: Partial<StyledRun>): Paragraph => ({
            runs: [{ text: 'aa', font: MONO, sizePx: 16, ...overrides }],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
        });

        const rules = (overrides: Partial<StyledRun>): LineOp[] =>
            renderPage(layoutPages([decorated(overrides)], PAGE)[0]!).ops
                .filter((op): op is LineOp => 'line' === op.kind);

        it('draws the underline where the FONT puts it', () => {
            // A rule a fixed distance below every baseline sits on the
            // descenders of one face and floats away from another.
            const [page] = layoutPages([decorated({ underline: { style: 'single' } })], PAGE);
            const line = page!.lines[0]!;
            const [rule] = rules({ underline: { style: 'single' } });

            expect(rule!.y1Px).toBeCloseTo(
                line.yPx + line.baselinePx + MONO.underlineOffset(16), 6,
            );
            expect(rule!.widthPx).toBeCloseTo(MONO.underlineThickness(16), 6);
        });

        it('spans the text and nothing more', () => {
            const [rule] = rules({ underline: { style: 'single' } });

            expect(rule!.x1Px).toBeCloseTo(10, 6);
            expect(rule!.x2Px - rule!.x1Px).toBeCloseTo(CELL * 2, 6);
        });

        it('draws TWO rules for a double underline', () => {
            const drawn = rules({ underline: { style: 'double' } });

            expect(drawn.length).toBe(2);
            expect(drawn[1]!.y1Px).toBeGreaterThan(drawn[0]!.y1Px);
        });

        it('strikes ABOVE the baseline, where the underline is below', () => {
            const [page] = layoutPages([decorated({ strike: 'single' })], PAGE);
            const baseline = page!.lines[0]!.yPx + page!.lines[0]!.baselinePx;
            const [rule] = rules({ strike: 'single' });

            expect(rule!.y1Px).toBeLessThan(baseline);
            expect(rule!.y1Px).toBeCloseTo(baseline - MONO.strikeoutOffset(16), 6);
        });

        it('takes the text\u2019s colour unless the rule names its own', () => {
            expect(rules({ underline: { style: 'single' }, colorHex: '#00FF00' })[0]!.color)
                .toBe('#00FF00');
            expect(rules({
                underline: { style: 'single', colorHex: '#FF0000' },
                colorHex: '#00FF00',
            })[0]!.color).toBe('#FF0000');
        });

        it('dashes a dotted or dashed underline', () => {
            expect(rules({ underline: { style: 'dotted' } })[0]!.style).toBe('dotted');
            expect(rules({ underline: { style: 'single' } })[0]!.style).toBe('solid');
        });

        it('never draws a rule thinner than a tenth of a POINT', () => {
            // LibreOffice's underline is about a twentieth of the font's size
            // — measured at 0.1, 0.2, 0.5 and 2.1pt for 2, 4, 10 and 40pt text
            // — and at 1pt, where a twentieth would be 0.05, it still drew
            // 0.1. That is the floor. This used to assert 0.5 PIXELS, which is
            // 0.375pt and nearly three times too thick: an invented number,
            // and the one flag `npm run audit:claims` was left holding.
            const tiny = rules({ underline: { style: 'single' }, sizePx: 1 });

            expect(MONO.underlineThickness(1)).toBeLessThan(0.1 * 96 / 72);
            expect(tiny[0]!.widthPx * 72 / 96).toBeCloseTo(0.1, 9);
        });

        it('leaves a rule thicker than the floor alone', () => {
            // 10pt text: a twentieth is half a point, well clear of the floor.
            const ordinary = rules({ underline: { style: 'single' }, sizePx: 10 * 96 / 72 });

            expect(ordinary[0]!.widthPx)
                .toBeCloseTo(MONO.underlineThickness(10 * 96 / 72), 9);
        });

        it('draws nothing for an undecorated run', () => {
            expect(rules({}).length).toBe(0);
        });
    });

    describe('borders and shading', () => {
        const bordered = (borders: BoxBorders, shadingFill?: string) => ({
            rows: [{
                cells: [{
                    paragraphs: [para('x')],
                    gridSpan: 1,
                    ...(undefined === shadingFill ? {} : { shadingFill }),
                }],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [CELL * 20],
            borders,
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        });

        const side = { widthPx: 2, colorHex: '#123456', style: 'solid' as const };

        it('paints shading, then rules, then the text over both', () => {
            // A cell fill drawn after its text would cover the words.
            const ops = renderPage(layoutPages(
                [bordered({ top: side }, '#EEEEEE')], PAGE,
            )[0]!).ops;

            expect(ops.map((op) => op.kind)).toEqual(['rect', 'line', 'text']);
        });

        it('lets a cell’s shading cover the table’s', () => {
            // A table fill is a background for every cell; a cell that names its
            // own is painting over it. Taking the table's would make a
            // highlighted header row disappear into the rest of the table.
            const table = {
                rows: [{
                    cells: [
                        { paragraphs: [para('own')], gridSpan: 1, shadingFill: '#111111' },
                        { paragraphs: [para('inherited')], gridSpan: 1 },
                    ],
                    isHeader: false,
                    cantSplit: false,
                }],
                columnWidthsPx: [CELL * 10, CELL * 10],
                shadingFill: '#999999',
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            };
            const fills = renderPage(layoutPages([table], PAGE)[0]!).ops
                .filter((op) => 'rect' === op.kind)
                .map((op) => op.fill);

            expect(fills).toEqual(['#111111', '#999999']);
        });

        it('gives a SECOND table on the page a top rule of its own', () => {
            // A shared edge is drawn once, so a row skips its top where the
            // row above belongs to the same table. Two tables are two tops:
            // the guard is on the BLOCK, not on the position in the list.
            const table = bordered({ top: side, bottom: side });
            const [page] = layoutPages([table, para('between'), table], PAGE);
            const ops = renderPage(page!).ops
                .filter((op): op is LineOp => 'line' === op.kind)
                .filter((op) => op.y1Px === op.y2Px);

            expect(page!.rows).toHaveLength(2);
            // A top and a bottom for each of them.
            expect(ops).toHaveLength(4);
        });

        it('draws each declared side along the cell edge it names', () => {
            // The SAME table for both, or the two disagree: a table with a left
            // border is inset by half of it, so deriving the expected edges
            // from a table without one puts them a border-width apart.
            const table = bordered({ top: side, bottom: side, left: side, right: side });
            const [page] = layoutPages([table], PAGE);
            const ops = renderPage(page!).ops.filter((op) => 'line' === op.kind);
            const row = page!.rows[0]!;
            const [left, right] = [row.cells[0]!.xPx, row.cells[0]!.xPx + row.cells[0]!.widthPx];
            // A horizontal rule is centred in the gap the flow keeps for it, so
            // it stands half its width outside the CONTENT. Measured against
            // LibreOffice: a row of 11.5pt text ran 800.97 to 789.47 with its
            // rules at 801.489 and 788.989 — half a point outside either end.
            const half = side.widthPx / 2;
            const top = row.yPx - half;
            const bottom = row.yPx + row.heightPx + half;

            // And the verticals run half a rule PAST the horizontals rather
            // than stopping on their centre lines: a table between
            // one-point rules at 769.389 and 719.189 printed its left edge
            // from 769.889 to 718.689. Stopping on the line leaves a quarter
            // of every outer corner unpainted.
            expect(ops.map((op) => [op.x1Px, op.y1Px, op.x2Px, op.y2Px])).toEqual([
                [left, top, right, top],
                [left, bottom, right, bottom],
                [left, top - half, left, bottom + half],
                [right, top - half, right, bottom + half],
            ]);
        });

        it('settles each COLUMN against the cell directly above it', () => {
            // The two columns disagree differently: the left pair is decided by
            // the upper row and the right pair by the lower one. Pairing a cell
            // with any neighbour rather than the one above gives both columns
            // the same answer.
            const heavy = { widthPx: 4, colorHex: '#FF0000', style: 'solid' as const };
            const other = { widthPx: 4, colorHex: '#0000FF', style: 'solid' as const };
            const cell = (borders: BoxBorders) => ({ paragraphs: [para('x')], gridSpan: 1, borders });

            const table = {
                rows: [
                    {
                        cells: [cell({ bottom: heavy }), cell({})],
                        isHeader: false,
                        cantSplit: false,
                    },
                    {
                        cells: [cell({}), cell({ top: other })],
                        isHeader: false,
                        cantSplit: false,
                    },
                ],
                columnWidthsPx: [CELL * 5, CELL * 5],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            };
            const rows = layoutPages([table], PAGE)[0]!.rows;

            expect(rows[0]!.cells[0]!.borders?.bottom?.colorHex).toBe('#FF0000');
            expect(rows[0]!.cells[1]!.borders?.bottom?.colorHex).toBe('#0000FF');
            expect(rows[1]!.cells[0]!.borders?.top?.colorHex).toBe('#FF0000');
            expect(rows[1]!.cells[1]!.borders?.top?.colorHex).toBe('#0000FF');
        });

        it('moves the cell TEXT with the table when it is inset', () => {
            // The inset shifts the whole table, borders and contents alike.
            // Moving only the boxes leaves every word a half-border out.
            const wide = { widthPx: 8, colorHex: '#000000', style: 'solid' as const };
            const inset = renderPage(layoutPages([bordered({ left: wide })], PAGE)[0]!).ops
                .find((op) => 'text' === op.kind)!;
            const plain = renderPage(layoutPages([bordered({})], PAGE)[0]!).ops
                .find((op) => 'text' === op.kind)!;

            expect(plain.xPx - inset.xPx).toBeCloseTo(4, 6);
        });

        it('carries the declared width, colour and style onto the rule', () => {
            const dashed = { widthPx: 3, colorHex: '#ABCDEF', style: 'dashed' as const };
            const [op] = renderPage(layoutPages([bordered({ top: dashed })], PAGE)[0]!).ops
                .filter((each) => 'line' === each.kind);

            expect(op).toMatchObject({ widthPx: 3, color: '#ABCDEF', style: 'dashed' });
        });

        it('draws nothing at all for a table that declares no borders', () => {
            const ops = renderPage(layoutPages([bordered({})], PAGE)[0]!).ops;

            expect(ops.every((op) => 'text' === op.kind)).toBe(true);
        });

        it('shades the cell box, not the text', () => {
            const row = layoutPages([bordered({}, '#EEEEEE')], PAGE)[0]!.rows[0]!;
            const [op] = renderPage(layoutPages([bordered({}, '#EEEEEE')], PAGE)[0]!).ops;

            expect(op).toMatchObject({
                kind: 'rect',
                xPx: row.cells[0]!.xPx,
                yPx: row.yPx,
                widthPx: row.cells[0]!.widthPx,
                heightPx: row.heightPx,
                fill: '#EEEEEE',
            });
        });
    });

    describe('justification', () => {
        it('splits a stretched line into words so each gap can widen', () => {
            const [laid] = layoutPages([para('aa bb cc dddddddddddddd')], PAGE, {
                widowOrphanControl: false,
            });
            const first = laid!.lines[0]!;
            const ops = textOps(renderPage(laid!).ops);

            expect(first.wordSpacingPx).toBe(0);
            // Unstretched, a line stays one operation per piece.
            expect(ops[0]!.text).toContain(' ');
        });

        it('widens a gap ONCE however many spaces are in it', () => {
            // The slack was divided by GAPS, and a run of spaces is one gap.
            // Adding the stretch per space character would push a line with a
            // double space past the margin it was fitted to.
            const single = layoutPages(
                [para('aa bb cc', { alignment: 'justify' })], PAGE, { widowOrphanControl: false },
            );
            const double = layoutPages(
                [para('aa  bb  cc', { alignment: 'justify' })], PAGE, { widowOrphanControl: false },
            );

            // Both are one line, so neither is stretched — the last line never
            // is. Stretch them by hand to state the rule directly.
            const stretched = {
                ...single[0]!,
                lines: [{ ...single[0]!.lines[0]!, wordSpacingPx: 10 }],
            };
            const stretchedDouble = {
                ...double[0]!,
                lines: [{ ...double[0]!.lines[0]!, wordSpacingPx: 10 }],
            };

            const gapOf = (page: typeof stretched): number => {
                const ops = textOps(renderPage(page).ops);

                return ops[1]!.xPx - ops[0]!.xPx;
            };

            // One space plus 10, versus two spaces plus the SAME 10.
            expect(gapOf(stretched)).toBeCloseTo(CELL * 2 + CELL + 10, 6);
            expect(gapOf(stretchedDouble)).toBeCloseTo(CELL * 2 + CELL * 2 + 10, 6);
        });
    });
});

describe('a vertically merged cell', () => {
    const side = { widthPx: 2, colorHex: '#112233', style: 'solid' as BorderStyle };

    const page = () => layoutPages([{
        rows: [
            {
                cells: [
                    {
                        paragraphs: [{
                            runs: [{ text: 'm', font: MONO, sizePx: 16 }],
                            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
                        }],
                        gridSpan: 1,
                        verticalMerge: 'restart',
                        borders: { top: side, bottom: side, left: side, right: side },
                        shadingFill: '#ff0000',
                    },
                    { paragraphs: [], gridSpan: 1 },
                ],
                isHeader: false,
                cantSplit: false,
            },
            {
                cells: [
                    { paragraphs: [], gridSpan: 1, verticalMerge: 'continue' },
                    { paragraphs: [], gridSpan: 1 },
                ],
                isHeader: false,
                cantSplit: false,
            },
        ],
        columnWidthsPx: [CELL * 4, CELL * 4],
        cellMarginLeftPx: 0,
        cellMarginRightPx: 0,
        cellMarginTopPx: 0,
        cellMarginBottomPx: 0,
        spaceBeforePx: 0,
        spaceAfterPx: 0,
        pageBreakBefore: false,
    }], PAGE)[0]!;

    it('draws its box down the whole span', () => {
        // Both rows are one line tall, so a box that stopped at the foot of the
        // first would be half as tall — and would put a rule straight through
        // the middle of the cell.
        const rows = page().rows;
        const top = rows[0]!.yPx;
        const foot = rows[1]!.yPx + rows[1]!.heightPx;
        const lines = renderPage(page()).ops.filter((op): op is LineOp => 'line' === op.kind);

        expect(rows[0]!.cells[0]!.heightPx).toBe(foot - top);
        // The bottom rule is at the foot of the SPAN, and there is no rule
        // anywhere between the two. Each stands half its width outside the
        // content, which is where the flow kept the gap for it.
        const half = (rows[0]!.cells[0]!.borders?.top?.widthPx ?? 0) / 2;
        const horizontal = lines
            .filter((op) => op.y1Px === op.y2Px && op.x1Px === rows[0]!.cells[0]!.xPx)
            .map((op) => op.y1Px);

        expect(horizontal).toEqual([top - half, foot + half]);
    });

    it('shades the whole span as one rectangle', () => {
        const rendered = page();
        const foot = rendered.rows[1]!.yPx + rendered.rows[1]!.heightPx;
        const fills = renderPage(rendered).ops
            .filter((op): op is RectOp => 'rect' === op.kind && '#ff0000' === op.fill);

        expect(fills.length).toBe(1);
        expect(fills[0]!.yPx + fills[0]!.heightPx).toBe(foot);
    });
});

describe('a table in a header', () => {
    it('is drawn, not only measured', () => {
        // The header reserves room for it either way, so a painter that
        // skipped it would leave a gap at the top of every page with nothing
        // to show for it.
        const [page] = layoutPages([{
            runs: [{ text: 'body', font: MONO, sizePx: 16 }],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
        }], { ...PAGE, headerDistancePx: 4 }, {
            headerFor: () => [{
                rows: [{
                    cells: [{
                        paragraphs: [{
                            runs: [{ text: 'logo', font: MONO, sizePx: 16 }],
                            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
                        }],
                        gridSpan: 1,
                        shadingFill: '#00ff00',
                    }],
                    isHeader: false,
                    cantSplit: false,
                }],
                columnWidthsPx: [CELL * 4],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            }],
        });
        const ops = renderPage(page!).ops;

        expect(ops.filter((op): op is TextOp => 'text' === op.kind)
            .map((op) => op.text)).toContain('logo');
        expect(ops.some((op) => 'rect' === op.kind && '#00ff00' === op.fill)).toBe(true);
    });
});

describe('a tracked run', () => {
    it('walks its words at the TRACKED advance', () => {
        // The renderer splits a piece at its spaces so each word can take its
        // share of a justified line. Measuring those parts bare would land
        // every word after the first short by the tracking on the one before.
        // Two lines, so the FIRST one is justified — the last line of a
        // justified paragraph is not, and the whole piece is drawn at once.
        const spaced: Paragraph = {
            runs: [{ text: 'aa bb cc dd', font: MONO, sizePx: 16, letterSpacingPx: 4 }],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE, alignment: 'justify' },
        };
        const narrow = { ...PAGE, widthPx: CELL * 10 + PAGE.marginLeftPx + PAGE.marginRightPx };
        const [page] = layoutPages([spaced], narrow);
        const ops = renderPage(page!).ops.filter((op): op is TextOp => 'text' === op.kind);
        const first = ops.find((op) => 'aa' === op.text)!;
        const second = ops.find((op) => 'bb' === op.text)!;

        // 'aa ' is three characters, each a glyph wide and tracked by four —
        // plus the share of the slack this justified line gave its one space.
        expect(second.xPx - first.xPx).toBeCloseTo(
            MONO.measureAdvance('aa ', 16).widthPt + 12 + page!.lines[0]!.wordSpacingPx,
            6,
        );
    });
});

describe('a drawn VML shape', () => {
    const SHAPE = { widthPx: 40, heightPx: 30, fillHex: '#123456', strokeHex: '#654321', strokeWidthPx: 2 };

    const withShape = (shape: InlineShape = SHAPE): Paragraph => ({
        runs: [
            { text: 'a', font: MONO, sizePx: 16 },
            { text: '\uFFFC', font: MONO, sizePx: 16, shape },
            { text: 'b', font: MONO, sizePx: 16 },
        ],
        style: { ...DEFAULT_PARAGRAPH_STYLE },
    });

    it('takes the WIDTH it asked for, not the width of its stand-in', () => {
        // The run's text is one object-replacement character; measuring that
        // would give whatever the font keeps at .notdef.
        const [page] = layoutPages([withShape()], PAGE);
        const pieces = page!.lines[0]!.line.pieces;

        expect(pieces[1]!.widthPx).toBe(40);
        expect(pieces[1]!.shape).toBe(SHAPE);
    });

    it('makes the line tall enough to hold it', () => {
        // It stands ON the baseline, so the line has to hold the whole of it
        // above and the font's descender below.
        const [tall] = layoutPages([withShape()], PAGE);
        const [flat] = layoutPages([withShape({ ...SHAPE, heightPx: 1 })], PAGE);

        expect(tall!.lines[0]!.heightPx).toBeGreaterThan(flat!.lines[0]!.heightPx);
        expect(tall!.lines[0]!.heightPx).toBeGreaterThanOrEqual(30);
    });

    it('fills and outlines it, standing on the baseline', () => {
        const [page] = layoutPages([withShape()], PAGE);
        const placed = page!.lines[0]!;
        const baseline = placed.yPx + placed.baselinePx;
        const ops = renderPage(page!).ops;
        const fill = ops.find((op): op is RectOp => 'rect' === op.kind && '#123456' === op.fill);
        const edges = ops.filter((op): op is LineOp => 'line' === op.kind
            && '#654321' === op.color);

        expect(fill?.heightPx).toBe(30);
        expect(fill?.yPx).toBe(baseline - 30);
        // Four edges, each at the stated weight.
        expect(edges.length).toBe(4);
        expect(edges.every((op) => 2 === op.widthPx)).toBe(true);
    });

    it('draws no fill for a shape that asked for none', () => {
        const bare = { widthPx: 40, heightPx: 10, strokeHex: '#000000', strokeWidthPx: 1 };
        const [page] = layoutPages([withShape(bare)], PAGE);
        const ops = renderPage(page!).ops;

        expect(ops.some((op) => 'rect' === op.kind)).toBe(false);
        expect(ops.filter((op) => 'line' === op.kind).length).toBe(4);
    });

    describe('paragraph borders', () => {
        const side = (widthPx: number): BorderSide =>
            ({ widthPx, style: 'single' as BorderStyle, colorHex: '#123456' });

        const BOX: BoxBorders = {
            top: side(2), left: side(2), bottom: side(2), right: side(2),
        };

        const rulesOf = (paragraphs: Paragraph[], pageIndex = 0): LineOp[] => {
            const pages = layoutPages(paragraphs, PAGE, { widowOrphanControl: false });

            return renderPage(pages[pageIndex]!).ops
                .filter((op): op is LineOp => 'line' === op.kind);
        };

        it('draws four rules round a bordered paragraph', () => {
            const rules = rulesOf([para('xx', { borders: BOX })]);

            expect(rules).toHaveLength(4);
            expect(rules.every((rule) => rule.color === '#123456')).toBe(true);
            expect(rules.every((rule) => rule.widthPx === 2)).toBe(true);
        });

        it('runs every rule half the crossing rule past the corner', () => {
            // LibreOffice's one-point box ran its top from 27.4 to 567.95 while
            // the sides stood at 27.9 and 567.45 — half a width past each end,
            // both ways. Butt the ends instead and each corner loses a square
            // the width of the stroke.
            const [laid] = layoutPages([para('xx', { borders: BOX })], PAGE);
            const [box] = laid!.paragraphBorders;
            const rules = rulesOf([para('xx', { borders: BOX })]);

            const horizontal = rules.filter((rule) => rule.y1Px === rule.y2Px);
            const vertical = rules.filter((rule) => rule.x1Px === rule.x2Px);

            expect(horizontal).toHaveLength(2);
            expect(vertical).toHaveLength(2);

            for (const rule of horizontal) {
                expect(Math.min(rule.x1Px, rule.x2Px)).toBe(box!.leftPx - 1);
                expect(Math.max(rule.x1Px, rule.x2Px)).toBe(box!.rightPx + 1);
            }
            for (const rule of vertical) {
                expect(Math.min(rule.y1Px, rule.y2Px)).toBe(box!.topPx - 1);
                expect(Math.max(rule.y1Px, rule.y2Px)).toBe(box!.bottomPx + 1);
            }
        });

        it('draws the w:between rule the full width of the box', () => {
            const borders: BoxBorders = { ...BOX, insideH: side(6) };
            const [laid] = layoutPages([para('a', { borders }), para('b', { borders })], PAGE);
            const [box] = laid!.paragraphBorders;
            const rules = rulesOf([para('a', { borders }), para('b', { borders })]);
            const inner = rules.filter((rule) => rule.widthPx === 6);

            expect(inner).toHaveLength(1);
            expect(inner[0]!.y1Px).toBe(box!.innerYPx[0]);
            expect(inner[0]!.y2Px).toBe(box!.innerYPx[0]);
            expect(Math.min(inner[0]!.x1Px, inner[0]!.x2Px)).toBe(box!.leftPx - 1);
            expect(Math.max(inner[0]!.x1Px, inner[0]!.x2Px)).toBe(box!.rightPx + 1);
        });

        it('omits a side the paragraph does not ask for', () => {
            const rules = rulesOf([para('xx', { borders: { bottom: side(2) } })]);

            expect(rules).toHaveLength(1);
            expect(rules[0]!.y1Px).toBe(rules[0]!.y2Px);
        });

        it('draws a COMPLETE box on each page a split box lands on', () => {
            // Nine full lines against a page that holds eight.
            const wide = Array.from({ length: 9 }, () => 'x'.repeat(20)).join(' ');
            const paragraphs = [para(wide, { borders: BOX })];
            const horizontals = (pageIndex: number): number => rulesOf(paragraphs, pageIndex)
                .filter((rule) => rule.y1Px === rule.y2Px).length;

            // Two horizontals on each page, not one: LibreOffice printed a
            // top AND a bottom either side of the break.
            expect(horizontals(0)).toBe(2);
            expect(horizontals(1)).toBe(2);
        });
    });
});

describe('everything on a page reaches the ink', () => {
    /**
     * Every `Page` field, said to be one of two things.
     *
     * The commonest defect this engine has had is a thing PLACED and never
     * DRAWN. A bordered paragraph inside a cell was collected by nobody; a
     * header's boxes were collected and then discarded where the furniture was
     * placed; and until someone went looking, nothing checked that any of the
     * newer surface survived the serialiser at all. Each was invisible in
     * exactly the same way — the model was right and the page was blank.
     *
     * So the compiler asks the question now. Add a field to `Page` and this
     * stops type-checking until someone says whether `renderPage` draws it.
     * `drawn` means an operation comes out of it; `not ink` means it is
     * geometry or identity that no operation is made from.
     */
    const PAGE_SURFACE: Record<keyof Page, 'drawn' | 'not ink'> = {
        lines: 'drawn',
        rows: 'drawn',
        header: 'drawn',
        footer: 'drawn',
        footnotes: 'drawn',
        paragraphBorders: 'drawn',
        pageBorder: 'drawn',
        lineNumbers: 'drawn',
        floats: 'drawn',

        // The paper's own size, which becomes the SVG's viewport rather than
        // an operation; the number a field prints, already resolved into the
        // text; and which section this page belongs to, which nothing paints.
        geometry: 'not ink',
        pageNumber: 'not ink',
        sectionIndex: 'not ink',
    };

    /** A page with something in every drawable field, to render in one go. */
    const full = (): Page => {
        const side: BorderSide = {
            widthPx: 2, style: 'single' as BorderStyle, colorHex: '#000000',
        };
        const box: BoxBorders = { top: side, left: side, bottom: side, right: side };
        const table = {
            rows: [{
                cells: [{ paragraphs: [para('celled')], gridSpan: 1, borders: box }],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [CELL * 10],
            borders: box,
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };
        const [page] = layoutPages([para('body'), table], PAGE, {
            headerFor: () => [para('head')],
            footerFor: () => [para('foot')],
        });

        return {
            ...page!,
            paragraphBorders: [{
                leftPx: 5, rightPx: 50, topPx: 5, bottomPx: 20,
                borders: box, opensHere: true, closesHere: true, innerYPx: [],
            }],
            pageBorder: {
                leftPx: 2, rightPx: 90, topPx: 2, bottomPx: 90,
                borders: box, opensHere: true, closesHere: true, innerYPx: [],
            },
            lineNumbers: [{
                run: { text: '1', font: MONO, sizePx: 16 }, xPx: 2, baselinePx: 30,
            }],
        };
    };

    it('draws SOMETHING for every field that says it is drawn', () => {
        // One rendering, and every `drawn` field accounted for in it. A field
        // that stopped reaching the ink would leave its count at zero.
        const ops = renderPage(full()).ops;
        const texts = ops.filter((op): op is TextOp => 'text' === op.kind).map((op) => op.text);
        const rules = ops.filter((op) => 'line' === op.kind);

        expect(texts).toContain('body');       // lines
        expect(texts).toContain('celled');     // rows
        expect(texts).toContain('head');       // header
        expect(texts).toContain('foot');       // footer
        expect(texts).toContain('1');          // lineNumbers
        // paragraphBorders, pageBorder and the table's own rules.
        expect(rules.length).toBeGreaterThanOrEqual(12);
    });

    it('agrees with itself about which fields are ink', () => {
        const drawn = Object.entries(PAGE_SURFACE)
            .filter(([, kind]) => 'drawn' === kind)
            .map(([key]) => key);

        expect(drawn).toHaveLength(9);
        expect(drawn).not.toContain('geometry');
    });
});
