import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import type { StyledRun } from '../../src/layout/line-breaker.js';
import {
    DEFAULT_PARAGRAPH_STYLE,
    DEFAULT_TAB_PX,
    FOOTNOTE_RULE_GAP_PX,
    FOOTNOTE_RULE_WIDTH_PX,
    layoutPages,
    layoutSections,
    placePageBorder,
    type Page,
    type PageGeometry,
    type Paragraph,
    type ParagraphStyle,
    type Section,
} from '../../src/layout/page-layout.js';
import type { Alignment } from '../../src/layout/alignment.js';
import type { BorderSide, BorderStyle, BoxBorders } from '../../src/layout/borders.js';
import type { FloatingImage } from '../../src/layout/float.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MONO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Regular.ttf'))));

/**
 * Two more faces, for the one question a single face cannot be asked.
 *
 * Carlito reaches FURTHER above the baseline than Liberation Mono and dives
 * LESS far below it, and Liberation Serif is the only one of the three whose
 * file carries a line gap. Sizes of one face cannot separate any of that.
 */
const CARLITO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'Carlito-Regular.ttf'))));
const SERIF = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationSerif-Regular.ttf'))));

const CELL = MONO.measureAdvance('x', 16).widthPt;
/** Fixed line height so page arithmetic is exact rather than font-dependent. */
const LINE = 20;

/** A page holding exactly 5 lines, 10 characters wide. */
const PAGE: PageGeometry = {
    widthPx: CELL * 10 + 20,
    heightPx: LINE * 5 + 20,
    marginTopPx: 10,
    marginRightPx: 10,
    marginBottomPx: 10,
    marginLeftPx: 10,
};

function para(text: string, style: Partial<ParagraphStyle> = {}): Paragraph {
    const runs: StyledRun[] = [{ text, font: MONO, sizePx: 16 }];

    return {
        runs,
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE, ...style },
    };
}

/** Words that each occupy one full line of the 10-character column. */
function lines(count: number): string {
    return Array.from({ length: count }, () => 'xxxxxxxxxx').join(' ');
}

function shape(pages: ReturnType<typeof layoutPages>): number[] {
    return pages.map((page) => page.lines.length);
}

describe('layoutPages', () => {
    describe('reflow', () => {
        it('continues on the next page instead of growing the first', () => {
            // The original complaint: content that overruns a page must move to
            // the next one. Five lines fit; the sixth starts page two.
            //
            // Widow control is off HERE so this states reflow and nothing else
            // — with it on the split is [4, 2], which is correct but is a
            // different rule, asserted separately below.
            expect(shape(layoutPages([para(lines(5))], PAGE))).toEqual([5]);
            expect(shape(layoutPages([para(lines(6))], PAGE, { widowOrphanControl: false })))
                .toEqual([5, 1]);
        });

        it('fills each page completely before starting another', () => {
            expect(shape(layoutPages([para(lines(12))], PAGE, { widowOrphanControl: false })))
                .toEqual([5, 5, 2]);
        });

        it('positions every line inside the page box', () => {
            const [page] = layoutPages([para(lines(5))], PAGE);

            for (const placed of page!.lines) {
                expect(placed.yPx).toBeGreaterThanOrEqual(PAGE.marginTopPx);
                expect(placed.yPx + placed.heightPx).toBeLessThanOrEqual(PAGE.heightPx - PAGE.marginBottomPx);
                expect(placed.xPx).toBe(PAGE.marginLeftPx);
            }
        });

        it('stops at the bottom margin, not at the paper edge', () => {
            // PAGE holds five lines whether or not its bottom margin counts, so
            // it cannot catch a layout that prints into the margin. This one
            // fits six without the margin and five with it.
            const geometry: PageGeometry = {
                ...PAGE,
                heightPx: LINE * 5 + 10 + 20,
                marginBottomPx: 20,
            };

            expect(shape(layoutPages([para(lines(6))], geometry, { widowOrphanControl: false })))
                .toEqual([5, 1]);
        });

        it('restarts at the top margin on each new page', () => {
            const pages = layoutPages([para(lines(6))], PAGE);

            expect(pages[1]!.lines[0]!.yPx).toBe(PAGE.marginTopPx);
        });
    });

    describe('paragraph spacing', () => {
        it('COLLAPSES space after against space before, taking the larger', () => {
            // This said "Word ADDS them; CSS collapses adjacent margins and
            // Word does not" — reasoned, never measured, and wrong about the
            // engine's own reference. Printed by LibreOffice with the
            // two set differently in BOTH orders, which is what tells `max`
            // from `sum`, from `after` alone, and from `before` alone:
            //
            //   10pt after, then 20pt before   gap 20.00
            //   20pt after, then 10pt before   gap 20.00
            //
            // Only the larger of the two explains both. Adding them put every
            // spaced paragraph further down its page than the file asks.
            //
            // OPEN: whether WORD agrees with LibreOffice here is not something
            // this repo can measure — there is no Word in the loop — and the
            // claim that it adds them is the one being overturned, so it should
            // not be taken on trust either. Anyone with Word to hand can settle
            // it with the same two pairs.
            const pages = layoutPages(
                [para('a', { spaceAfterPx: 6 }), para('b', { spaceBeforePx: 4 })],
                PAGE,
            );

            const [first, second] = pages[0]!.lines;
            expect(second!.yPx - (first!.yPx + first!.heightPx)).toBe(6);
        });

        it('collapses inside a table CELL as well as on the page', () => {
            // A cell stacks paragraphs the way the page does, through
            // `stackBlocks`, and the collapsing rule belongs to both.
            // No printed page decides this one: what it says is that the two
            // stacking loops do not drift apart, which they did until a
            // mutation collapsing only the page's went unnoticed.
            const cell = {
                paragraphs: [para('a', { spaceAfterPx: 6 }), para('b', { spaceBeforePx: 4 })],
                gridSpan: 1,
            };
            const table = {
                rows: [{ cells: [cell], isHeader: false, cantSplit: false }],
                columnWidthsPx: [CELL * 10],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            };
            const [first, second] = layoutPages([table], PAGE)[0]!.rows[0]!.cells[0]!.lines;

            expect(second!.yPx - (first!.yPx + first!.heightPx)).toBe(6);
        });

        it('does not collapse a paragraph’s space against a TABLE above it', () => {
            // A DECISION, not a measurement: a table's own spacing and a
            // paragraph's are separate quantities in the file, so the space
            // below a table is not spent against the space above the next
            // paragraph. Nothing printed says so — the fixture the collapsing
            // rule was measured from has no table between two spaced
            // paragraphs — and it is asserted here so that changing it is a
            // choice somebody makes rather than a side effect they do not
            // notice.
            const table = {
                rows: [{
                    cells: [{ paragraphs: [para('x')], gridSpan: 1 }],
                    isHeader: false,
                    cantSplit: false,
                }],
                columnWidthsPx: [CELL * 10],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 6,
                pageBreakBefore: false,
            };
            const [page] = layoutPages(
                [para('a', { spaceAfterPx: 6 }), table, para('b', { spaceBeforePx: 4 })],
                PAGE,
            );
            const [first, second] = page!.lines;
            const row = page!.rows[0]!;

            expect(second!.yPx - (row.yPx + row.heightPx)).toBe(10);
            expect(first).toBeDefined();
        });

        it('takes the larger the other way round too', () => {
            const pages = layoutPages(
                [para('a', { spaceAfterPx: 4 }), para('b', { spaceBeforePx: 6 })],
                PAGE,
            );

            const [first, second] = pages[0]!.lines;
            expect(second!.yPx - (first!.yPx + first!.heightPx)).toBe(6);
        });

        it('applies space before at the top of a page', () => {
            // Word applies it; only the legacy compatibility option
            // suppressSpBfAfterPgBrk drops it. Dropping it unconditionally
            // would raise the first line of every page by the paragraph's own
            // space-before, which moves every break after it.
            const [page] = layoutPages([para('a', { spaceBeforePx: 8 })], PAGE);

            expect(page!.lines[0]!.yPx).toBe(PAGE.marginTopPx + 8);
        });
    });

    describe('indents', () => {
        it('indents the first line only, and narrows it to match', () => {
            // A first-line indent that moved the text without narrowing the
            // column would let the first line run past the right margin.
            //
            // The words are sized so the NARROWING decides where the break
            // falls: seven cells hold 'aa bb', the full ten would hold
            // 'aa bb cc'. An assertion that passes either way would not see the
            // difference, so the width is pinned exactly.
            const pages = layoutPages(
                [para('aa bb cc dd', { indentFirstLinePx: CELL * 3 })],
                PAGE,
            );

            const [first, second] = pages[0]!.lines;
            expect(first!.xPx).toBe(PAGE.marginLeftPx + CELL * 3);
            expect(second!.xPx).toBe(PAGE.marginLeftPx);
            expect(first!.line.widthPx).toBeCloseTo(CELL * 5, 6);
            expect(first!.line.overflows).toBe(false);
        });

        it('applies left and right indents to every line', () => {
            const pages = layoutPages(
                [para('aaa bbb ccc', { indentLeftPx: CELL * 2, indentRightPx: CELL * 2 })],
                PAGE,
            );

            expect(pages[0]!.lines.length).toBeGreaterThan(1);
            for (const placed of pages[0]!.lines) {
                expect(placed.xPx).toBe(PAGE.marginLeftPx + CELL * 2);
                expect(placed.line.widthPx).toBeLessThanOrEqual(CELL * 6);
                expect(placed.line.overflows).toBe(false);
            }
        });
    });

    describe('explicit page breaks', () => {
        it('starts a new page even with room to spare', () => {
            const pages = layoutPages([para('a'), para('b', { pageBreakBefore: true })], PAGE);

            expect(shape(pages)).toEqual([1, 1]);
        });

        it('does not open a blank page when already at the top', () => {
            // Otherwise a document whose first paragraph carries the flag opens
            // with an empty sheet.
            const pages = layoutPages([para('a', { pageBreakBefore: true })], PAGE);

            expect(shape(pages)).toEqual([1]);
        });
    });

    describe('widow and orphan control', () => {
        it('never strands the first line of a paragraph at the foot of a page', () => {
            // Four lines used, one slot left, and a 3-line paragraph to place.
            // Placing one line there orphans it, so the whole paragraph moves.
            const pages = layoutPages([para(lines(4)), para(lines(3))], PAGE);

            expect(shape(pages)).toEqual([4, 3]);
        });

        it('never strands the last line of a paragraph at the head of a page', () => {
            // A 6-line paragraph on a 5-line page would leave line 6 alone, so
            // line 5 is pushed across to keep it company.
            expect(shape(layoutPages([para(lines(6))], PAGE))).toEqual([4, 2]);
        });

        it('can be switched off, and then splits wherever it likes', () => {
            expect(shape(layoutPages([para(lines(6))], PAGE, { widowOrphanControl: false })))
                .toEqual([5, 1]);
        });

        it('keeps a marked paragraph whole', () => {
            // "Keep lines together" — what a heading or caption uses.
            const pages = layoutPages(
                [para(lines(3)), para(lines(3), { keepLinesTogether: true })],
                PAGE,
            );

            expect(shape(pages)).toEqual([3, 3]);
        });

        it('still places a paragraph too tall for any page', () => {
            // Nothing can be done for it, and refusing to split would loop
            // forever. It must be laid out, not hang the engine.
            expect(shape(layoutPages([para(lines(8), { keepLinesTogether: true })], PAGE)))
                .toEqual([5, 3]);
        });
    });

    describe('line provenance', () => {
        it('records which paragraph and which line each placed line came from', () => {
            // Editing and hit-testing both need to map a point on the page back
            // to a position in the document; without this the layout is a
            // picture rather than a model.
            const pages = layoutPages([para(lines(2)), para('b')], PAGE);
            const placed = pages[0]!.lines;

            expect(placed.map((p) => [p.paragraphIndex, p.lineIndex])).toEqual([[0, 0], [0, 1], [1, 0]]);
        });

        it('numbers lines within the PARAGRAPH, not within the page', () => {
            // A paragraph that spans a break is the only thing that tells the
            // two apart. Page-relative numbering would make the first line of
            // every page claim to be line 0 and send the caret to the wrong
            // place on every hit-test after a break.
            const pages = layoutPages([para(lines(6))], PAGE);

            expect(pages[1]!.lines.map((p) => p.lineIndex)).toEqual([4, 5]);
        });
    });

    describe('line height', () => {
        it('falls back to the font\'s natural single spacing', () => {
            // Every other test here pins lineHeightPx so the page arithmetic is
            // exact. Without this one, the fallback that real documents use
            // would never run.
            const runs: StyledRun[] = [{ text: 'a', font: MONO, sizePx: 16 }];
            const [page] = layoutPages(
                [{ runs, style: { ...DEFAULT_PARAGRAPH_STYLE } }],
                PAGE,
            );

            expect(page!.lines[0]!.heightPx).toBeCloseTo(MONO.naturalLineHeight(16), 9);
        });

        it('honours an EXACT line height, overriding the font', () => {
            // Word's "Exactly" line spacing. The font's own height here is
            // 18.125px and the paragraph asks for 20 — but both fit five lines
            // on this page, so only the measured height and the distance
            // between baselines can tell them apart.
            expect(MONO.naturalLineHeight(16)).not.toBeCloseTo(LINE, 3);

            const [page] = layoutPages([para(lines(2))], PAGE);

            expect(page!.lines[0]!.heightPx).toBe(LINE);
            expect(page!.lines[1]!.yPx - page!.lines[0]!.yPx).toBe(LINE);
        });

        it('takes the TALLEST run on the line, not the first', () => {
            // One larger word makes the whole line taller. Sizing a line from
            // its first run would let a mid-sentence heading overlap the line
            // below.
            //
            // Both runs have to LAND on one line for this to be asked at all:
            // 'ab ' is three cells at 16px and 'BIG' six at 32, which fits the
            // ten this page is wide. The fixture said 'small ' at first and
            // spilled 'BIG' onto a second line, so what it really measured was
            // the paragraph — the one thing its name says it is not.
            const runs: StyledRun[] = [
                { text: 'ab ', font: MONO, sizePx: 16 },
                { text: 'BIG', font: MONO, sizePx: 32 },
            ];
            const [page] = layoutPages(
                [{ runs, style: { ...DEFAULT_PARAGRAPH_STYLE } }],
                PAGE,
            );

            expect(page!.lines).toHaveLength(1);
            expect(page!.lines[0]!.heightPx).toBeCloseTo(MONO.naturalLineHeight(32), 9);
        });

        it('leaves the OTHER lines alone: a box belongs to its line', () => {
            // The same two runs, too wide to share a line, so the big word
            // lands on the second. Measured against LibreOffice: a
            // paragraph whose first line carries a 36pt picture printed that
            // line 38.10 tall and every line after it 11.50, and moving the
            // picture down the paragraph moved the tall line with it. A height
            // measured per PARAGRAPH gives both lines the big one and spaces
            // ordinary text by whatever the paragraph's largest thing is.
            const runs: StyledRun[] = [
                { text: 'small ', font: MONO, sizePx: 16 },
                { text: 'BIG', font: MONO, sizePx: 32 },
            ];
            const [page] = layoutPages(
                [{ runs, style: { ...DEFAULT_PARAGRAPH_STYLE } }],
                PAGE,
            );

            expect(page!.lines).toHaveLength(2);
            expect(page!.lines[0]!.heightPx).toBeCloseTo(MONO.naturalLineHeight(16), 9);
            expect(page!.lines[1]!.heightPx).toBeCloseTo(MONO.naturalLineHeight(32), 9);
            // And the second line starts one SHORT line down, not one tall one.
            expect(page!.lines[1]!.yPx - page!.lines[0]!.yPx)
                .toBeCloseTo(MONO.naturalLineHeight(16), 9);
        });
    });

    describe("a line box holding two FACES", () => {
        /** Room for any of these on ONE line, so the box is the only variable. */
        const ROOMY: PageGeometry = {
            widthPx: 2000, heightPx: 2000,
            marginTopPx: 0, marginRightPx: 0, marginBottomPx: 0, marginLeftPx: 0,
        };
        const SIZE = 22;

        /** The height of the single line two runs share. */
        function boxOf(first: TrueTypeFont, second: TrueTypeFont): number {
            const runs: StyledRun[] = [
                { text: 'ab ', font: first, sizePx: SIZE },
                { text: 'cd', font: second, sizePx: SIZE },
            ];
            const [page] = layoutPages([{ runs, style: { ...DEFAULT_PARAGRAPH_STYLE } }], ROOMY);

            expect(page!.lines).toHaveLength(1);

            return page!.lines[0]!.heightPx;
        }

        it('reaches as far above as one face asks and as far below as the other', () => {
            //  MEASURED against LibreOffice at 22pt: a line holding Carlito
            // and Liberation Mono stepped **27.60**. Carlito reaches
            // 20.947 above the baseline and Liberation Mono 6.606 below it,
            // and neither face asks for both.
            //
            // The tallest WHOLE line among the runs is Carlito's 26.855, which
            // is what this used to answer -- 0.75pt short of every line, one
            // line in every thirty-six. A single face can never show it: its
            // tallest total is always also its deepest descender.
            expect(boxOf(CARLITO, MONO)).toBeCloseTo(27.5537, 4);
            expect(boxOf(CARLITO, MONO)).not.toBeCloseTo(CARLITO.naturalLineHeight(SIZE), 2);
            // Which side the face is on is not a question a line box asks.
            expect(boxOf(MONO, CARLITO)).toBeCloseTo(boxOf(CARLITO, MONO), 9);
        });

        it('keeps a face\'s LINE GAP with that face, rather than maxing it apart', () => {
            // Liberation Serif carries a 0.94pt gap at this size and Carlito
            // none, but Carlito wins both sides outright -- so the gap never
            // reaches the line. MEASURED: 26.85, Carlito's own line.
            //
            // A gap maxed on its own axis would have made this 27.79.
            expect(boxOf(CARLITO, SERIF)).toBeCloseTo(CARLITO.naturalLineHeight(SIZE), 9);
            expect(boxOf(CARLITO, SERIF)).not.toBeCloseTo(27.79, 1);
        });

        it('counts a gap where its own face reaches highest', () => {
            // The same serif beside Liberation Mono, which it out-reaches: now
            // the gap is part of the tallest ascent and the mono supplies the
            // descender. MEASURED: 27.20 against the 27.146 computed here --
            // LibreOffice's own rounding, and a twentieth of what the old rule
            // was out by.
            expect(boxOf(SERIF, MONO)).toBeCloseTo(27.1455, 4);
        });

        it('is unchanged where one face wins both sides', () => {
            // The control the measurement needed: Carlito reaches further above
            // AND further below than Caladea, so both rules answer the same
            // number and the printed step (26.85) cannot choose between them.
            // Liberation Serif stands in for Caladea here, having the same
            // shape of answer.
            expect(boxOf(CARLITO, SERIF)).toBeCloseTo(boxOf(CARLITO, CARLITO), 9);
        });
    });

    describe('tab stops', () => {
        /** A paragraph with one tab in it, laid out on a very wide page. */
        const wide = { ...PAGE, widthPx: CELL * 200, heightPx: 10000 };

        const widthOf = (style: Partial<ParagraphStyle>): number => {
            const paragraph: Paragraph = {
                runs: [{ text: 'a\tb', font: MONO, sizePx: 16 }],
                style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE, ...style },
            };

            return layoutPages([paragraph], wide)[0]!.lines[0]!.line.widthPx;
        };

        it('carries the paragraph\'s own stops into the line breaking', () => {
            // Reading w:tabs into the model is not the same as USING it: the
            // stops have to reach the breaker or the column is the default one.
            expect(widthOf({ tabStops: [{ positionPx: 300, align: 'left' }] })).toBeCloseTo(300 + CELL, 6);
            expect(widthOf({})).toBeCloseTo(DEFAULT_TAB_PX + CELL, 6);
        });

        it('measures stops from the paragraph, not from an indented first line', () => {
            // The first line starts 100px in, so the stop at 300 is 200 away —
            // and the tabbed text lands in the same column as it would on any
            // other line of the paragraph.
            const indented = widthOf({ tabStops: [{ positionPx: 300, align: 'left' }], indentFirstLinePx: 100 });

            expect(indented).toBeCloseTo(200 + CELL, 6);
        });
    });

    describe('baselines', () => {
        const DESCENT = MONO.descent(16);
        const NATURAL = MONO.naturalLineHeight(16);

        it('gives every placed line the baseline for its rule', () => {
            // LINE is 20 and the font's natural height is not, so `auto` and
            // `atLeast` disagree here and the test can see which was used.
            const at = (style: Partial<ParagraphStyle>): number =>
                layoutPages([para('xx', style)], PAGE)[0]!.lines[0]!.baselinePx;

            expect(at({ lineRule: 'auto' })).toBeCloseTo(NATURAL - DESCENT, 6);
            expect(at({ lineRule: 'atLeast' })).toBeCloseTo(LINE - DESCENT, 6);
            expect(at({ lineRule: 'exact' })).toBeCloseTo(LINE * 0.8, 6);
        });

        it('treats an unstated rule as auto', () => {
            const [page] = layoutPages([para('xx')], PAGE);

            expect(page!.lines[0]!.baselinePx).toBeCloseTo(NATURAL - DESCENT, 6);
        });

        it('takes the descent from the DEEPEST run on the line', () => {
            // A line is as deep as its deepest run, for the same reason it is as
            // tall as its tallest: sizing from another run clips the descender.
            const mixed: Paragraph = {
                runs: [
                    { text: 'aa', font: MONO, sizePx: 16 },
                    { text: 'bb', font: MONO, sizePx: 32 },
                ],
                style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: 60, lineRule: 'atLeast' },
            };

            expect(layoutPages([mixed], { ...PAGE, heightPx: 400 })[0]!.lines[0]!.baselinePx)
                .toBeCloseTo(60 - MONO.descent(32), 6);
        });

        it('gives a line inside a table CELL its baseline too', () => {
            const table = {
                rows: [{
                    cells: [{ paragraphs: [para('xx', { lineRule: 'exact' })], gridSpan: 1 }],
                    isHeader: false,
                    cantSplit: false,
                }],
                columnWidthsPx: [CELL * 10],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            };

            expect(layoutPages([table], PAGE)[0]!.rows[0]!.cells[0]!.lines[0]!.baselinePx)
                .toBeCloseTo(LINE * 0.8, 6);
        });
    });

    describe('alignment', () => {
        const widthOf = (alignment: Alignment): number =>
            layoutPages([para('xx', { alignment })], PAGE)[0]!.lines[0]!.xPx;

        it('offsets a placed line by its alignment', () => {
            // The column is ten cells wide inside a 10px left margin; two
            // characters leave eight cells of slack.
            expect(widthOf('left')).toBeCloseTo(10, 6);
            expect(widthOf('center')).toBeCloseTo(10 + CELL * 4, 6);
            expect(widthOf('right')).toBeCloseTo(10 + CELL * 8, 6);
        });

        it('carries the document\u2019s decimal symbol through to the line', () => {
            // The whole chain: a comma-decimal document ranges its numbers on
            // the comma, and reading the symbol without passing it on ranges
            // them on nothing at all.
            const decimal = (decimalSymbol?: string): number => {
                const paragraph = para('a\t12,34', {
                    tabStops: [{ positionPx: CELL * 8, align: 'decimal' }],
                    ...(undefined === decimalSymbol ? {} : { decimalSymbol }),
                });

                return layoutPages([paragraph], { ...PAGE, widthPx: CELL * 40 })[0]!
                    .lines[0]!.line.pieces.slice(0, 2)
                    .reduce((sum, piece) => sum + piece.widthPx, 0);
            };

            // Two digits before the comma, so the comma lands on the stop.
            expect(decimal(',')).toBeCloseTo(CELL * 6, 6);
            // Without it there is no separator to find, so the whole number is
            // ranged and the text starts five cells earlier.
            expect(decimal()).toBeCloseTo(CELL * 3, 6);
        });

        it('measures the FIRST line against its own narrower box', () => {
            // A first-line indent takes room away from that line only, so a
            // centred first line is centred in what is left — not in the column.
            const [page] = layoutPages(
                [para('xxxx yyyy', { alignment: 'center', indentFirstLinePx: CELL * 2 })],
                PAGE,
            );
            const [first, second] = page!.lines;

            expect(page!.lines.length).toBe(2);
            // Line one has 2 cells of indent and 8 cells left; 4 are used, so it
            // is offset by 2 ON TOP of the indent.
            expect(first!.xPx).toBeCloseTo(10 + CELL * 2 + CELL * 2, 6);
            // Line two has the whole 10 cells and 4 used, so it is offset by 3 —
            // and ends up further LEFT than the line above it, which is what
            // centring in a narrower box means.
            expect(second!.xPx).toBeCloseTo(10 + CELL * 3, 6);
            expect(second!.xPx).toBeLessThan(first!.xPx);
        });

        it('carries the justification stretch onto the line', () => {
            const [page] = layoutPages([para('aa bb ccccc dd', { alignment: 'justify' })], PAGE);
            const [first, last] = page!.lines;

            expect(first!.wordSpacingPx).toBeGreaterThan(0);
            // The last line of the paragraph is never stretched.
            expect(last!.wordSpacingPx).toBe(0);
        });

        it('aligns a paragraph INSIDE a table cell the same way', () => {
            // The cells go through their own stacking code, so an alignment
            // honoured on the page and ignored in a cell is exactly the drift
            // the shared helper exists to prevent.
            const cell = (alignment: Alignment) => ({
                rows: [{
                    cells: [{ paragraphs: [para('xx', { alignment })], gridSpan: 1 }],
                    isHeader: false,
                    cantSplit: false,
                }],
                columnWidthsPx: [CELL * 10],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            });

            const xOf = (alignment: Alignment): number =>
                layoutPages([cell(alignment)], PAGE)[0]!.rows[0]!.cells[0]!.lines[0]!.xPx;

            // Ten cells of column, two used: centred is four in, right is eight.
            expect(xOf('left')).toBeCloseTo(10, 6);
            expect(xOf('center')).toBeCloseTo(10 + CELL * 4, 6);
            expect(xOf('right')).toBeCloseTo(10 + CELL * 8, 6);
        });

        it('aligns the page FURNITURE the same way', () => {
            // A centred header is the commonest aligned thing in a document,
            // and it stacks through the same shared path a cell does.
            const [page] = layoutPages([para('body')], { ...PAGE, headerDistancePx: 4 }, {
                headerFor: () => [para('xx', { alignment: 'center' })],
            });

            expect(page!.header?.lines[0]!.xPx).toBeCloseTo(10 + CELL * 4, 6);
        });

        it('gives an unjustified line no word spacing at all', () => {
            expect(layoutPages([para('aa bb', { alignment: 'right' })], PAGE)[0]!.lines[0]!
                .wordSpacingPx).toBe(0);
        });
    });

    describe('floating pictures', () => {
        const picture = (widthPx: number, heightPx: number) => ({
            content: { bytes: new Uint8Array([1]), contentType: 'image/png' },
            widthPx,
            heightPx,
        });

        const anchored = (overrides: Partial<FloatingImage> = {}): FloatingImage => ({
            image: picture(CELL * 4, LINE * 2),
            horizontal: { relativeTo: 'column', offsetPx: 0 },
            vertical: { relativeTo: 'paragraph', offsetPx: 0 },
            wrap: 'square',
            marginTopPx: 0,
            marginBottomPx: 0,
            marginLeftPx: 0,
            marginRightPx: 0,
            behindText: false,
            ...overrides,
        });

        const withFloat = (float: FloatingImage, text = lines(6)): Paragraph => ({
            ...para(text),
            floats: [float],
        });

        it('places a float against the top of the paragraph it is anchored to', () => {
            const [page] = layoutPages([para('first'), withFloat(anchored())], PAGE, {
                widowOrphanControl: false,
            });

            // One line of text above it, so the float starts a line down.
            expect(page!.floats.length).toBe(1);
            expect(page!.floats[0]!.yPx).toBeCloseTo(10 + LINE, 6);
            expect(page!.floats[0]!.xPx).toBeCloseTo(10, 6);
        });

        it('measures an offset from the writing area, not the paper', () => {
            const [page] = layoutPages([withFloat(anchored({
                horizontal: { relativeTo: 'column', offsetPx: 30 },
            }))], PAGE, { widowOrphanControl: false });

            expect(page!.floats[0]!.xPx).toBeCloseTo(10 + 30, 6);
        });

        it('measures a PAGE-relative offset from the paper', () => {
            const [page] = layoutPages([withFloat(anchored({
                horizontal: { relativeTo: 'page', offsetPx: 30 },
            }))], PAGE, { widowOrphanControl: false });

            expect(page!.floats[0]!.xPx).toBeCloseTo(30, 6);
        });

        it('honours an alignment instead of an offset', () => {
            const [page] = layoutPages([withFloat(anchored({
                horizontal: { relativeTo: 'column', offsetPx: 999, align: 'right' },
            }))], PAGE, { widowOrphanControl: false });

            // Flush with the right of the writing area, and the offset ignored.
            expect(page!.floats[0]!.xPx).toBeCloseTo(10 + CELL * 10 - CELL * 4, 6);
        });

        it('aligns within what the float LEFT, not within the column', () => {
            // A centred line beside a float is centred in the six cells it
            // actually has. Centring it in the ten the column has would put it
            // two cells left of where it belongs — under the float.
            const centred: Paragraph = {
                ...para('xx', { alignment: 'center' }),
                floats: [anchored()],
            };
            const [page] = layoutPages([centred], PAGE, { widowOrphanControl: false });

            // Box starts at 4 cells in and is 6 wide; two characters centre two
            // cells into it.
            expect(page!.lines[0]!.xPx).toBeCloseTo(10 + CELL * 4 + CELL * 2, 6);
        });

        it('keeps the clear space on the LEFT of a float too', () => {
            // The float sits at the right, so the text stays left of it — and
            // has to stop short of the space it keeps clear.
            const spaced = anchored({
                horizontal: { relativeTo: 'column', offsetPx: CELL * 6 },
                marginLeftPx: CELL,
            });
            // Short words: a ten-character word does not FIT five cells and
            // is placed whole regardless, which would hide the narrowing.
            const [page] = layoutPages([{ ...para('aa bb cc dd'), floats: [spaced] }], PAGE, {
                widowOrphanControl: false,
            });

            // Six cells of column less one of clear space.
            expect(page!.lines[0]!.line.widthPx).toBeLessThanOrEqual(CELL * 5);
        });

        it('narrows the lines beside it and gives the column back below', () => {
            const [page] = layoutPages([withFloat(anchored())], PAGE, {
                widowOrphanControl: false,
            });
            const xs = page!.lines.map((line) => line.xPx);

            // The float is 4 cells wide and 2 lines tall at the left edge, so
            // the first two lines start beyond it and the rest do not.
            expect(xs[0]).toBeCloseTo(10 + CELL * 4, 6);
            expect(xs[1]).toBeCloseTo(10 + CELL * 4, 6);
            expect(xs[2]).toBeCloseTo(10, 6);
        });

        it('breaks the narrowed lines SHORTER, not just moves them', () => {
            // A line beside the float has six cells rather than ten, so it holds
            // fewer words — moving it without re-breaking would run it off the
            // page.
            const [page] = layoutPages([withFloat(anchored(), 'aa bb cc dd ee ff gg hh')], PAGE, {
                widowOrphanControl: false,
            });

            expect(page!.lines[0]!.line.widthPx).toBeLessThanOrEqual(CELL * 6);
        });

        it('leaves the text alone for a float that wraps NONE', () => {
            const [page] = layoutPages([withFloat(anchored({ wrap: 'none' }))], PAGE, {
                widowOrphanControl: false,
            });

            expect(page!.lines[0]!.xPx).toBeCloseTo(10, 6);
            // It is still placed, and still drawn.
            expect(page!.floats.length).toBe(1);
        });

        it('keeps the float on the page its anchor starts', () => {
            // The float belongs to the paragraph, and the paragraph starts here.
            const pages = layoutPages([para(lines(4)), withFloat(anchored())], PAGE, {
                widowOrphanControl: false,
            });

            expect(pages[0]!.floats.length).toBe(1);
            expect(pages.slice(1).every((page) => 0 === page.floats.length)).toBe(true);
        });
    });

    describe('page furniture', () => {
        /** Room for the furniture: a header 14px down and a footer 12px up. */
        const PAPER: PageGeometry = { ...PAGE, headerDistancePx: 14, footerDistancePx: 12 };

        it('hangs the header from its distance off the TOP edge', () => {
            const [page] = layoutPages([para('body')], PAPER, {
                headerFor: () => [para('h')],
            });

            expect(page!.header?.topPx).toBe(14);
            expect(page!.header?.lines[0]!.yPx).toBe(14);
            expect(page!.header?.heightPx).toBe(LINE);
        });

        it('stands the footer ON its distance off the BOTTOM edge', () => {
            // The one rule a shared implementation would get wrong. A footer is
            // anchored by its BOTTOM, so a taller one starts higher up and both
            // end in the same place; anchoring it by its top like the header
            // would run a two-line footer off the paper.
            const of = (...text: string[]) => layoutPages([para('body')], PAPER, {
                footerFor: () => text.map((line) => para(line)),
            })[0]!.footer!;

            const one = of('f');
            const two = of('f', 'f');

            expect(two.heightPx).toBe(LINE * 2);
            expect(one.topPx).toBe(PAPER.heightPx - 12 - LINE);
            expect(two.topPx).toBe(PAPER.heightPx - 12 - LINE * 2);
            expect(one.topPx + one.heightPx).toBe(two.topPx + two.heightPx);
        });

        it('places and reports a TABLE in a header', () => {
            // A logo beside an address is a table. Measuring it and then
            // dropping it would leave the header the right height with nothing
            // in it, which is worse than either.
            const [page] = layoutPages([para('body')], PAPER, {
                headerFor: () => [{
                    rows: [{
                        cells: [{ paragraphs: [para('logo')], gridSpan: 1 }],
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
            const rows = page!.header?.rows ?? [];

            expect(rows.length).toBe(1);
            // In PAGE coordinates, like the header's lines are.
            expect(rows[0]!.yPx).toBe(14);
            expect(rows[0]!.cells[0]!.xPx).toBe(PAPER.marginLeftPx);
            expect(rows[0]!.cells[0]!.lines[0]!.line.pieces[0]!.text).toBe('logo');
        });

        it('reports furniture in PAGE coordinates, left margin included', () => {
            const [page] = layoutPages([para('body')], PAPER, {
                headerFor: () => [para('h', { indentLeftPx: 7 })],
            });

            expect(page!.header?.lines[0]!.xPx).toBe(PAPER.marginLeftPx + 7);
        });

        it('repeats the furniture on every page', () => {
            const pages = layoutPages([para(lines(12))], PAPER, {
                widowOrphanControl: false,
                headerFor: () => [para('h')],
                footerFor: () => [para('f')],
            });

            expect(pages.length).toBe(3);
            expect(pages.every((page) => undefined !== page.header && undefined !== page.footer))
                .toBe(true);
        });

        it('leaves a page with no furniture without any', () => {
            const [page] = layoutPages([para('body')], PAPER, {
                // Declared, but this page is not one it applies to.
                headerFor: (pageIndex) => 1 === pageIndex ? [para('h')] : undefined,
                footerFor: () => [],
            });

            expect(page!.header).toBeUndefined();
            expect(page!.footer).toBeUndefined();
        });

        it('wraps furniture in the WRITING width, not the width of the paper', () => {
            // Twelve characters: they fit across the paper, which is twelve and
            // a bit wide, and they do not fit the ten-cell column between the
            // margins. Measuring against the sheet gives a one-line header and
            // starts the body a line too high.
            const [page] = layoutPages([para('body')], PAPER, {
                headerFor: () => [para('xxxxxx xxxxx')],
            });

            expect(CELL * 12).toBeLessThan(PAPER.widthPx);
            expect(page!.header?.lines.length).toBe(2);
            expect(page!.header?.heightPx).toBe(LINE * 2);
        });

        it('numbers pages from one', () => {
            const pages = layoutPages([para(lines(12))], PAGE, { widowOrphanControl: false });

            expect(pages.map((page) => page.pageNumber)).toEqual([1, 2, 3]);
        });

        it('asks for furniture by the page it is drawing', () => {
            const asked: [number, number][] = [];
            layoutPages([para(lines(12))], PAPER, {
                widowOrphanControl: false,
                headerFor: (pageIndex, pageNumber) => {
                    asked.push([pageIndex, pageNumber]);

                    return [para('h')];
                },
            });

            expect(asked).toEqual([[0, 1], [1, 2], [2, 3]]);
        });
    });

    describe('a writing area that varies by page', () => {
        it('asks for the box of EACH page, not once for the document', () => {
            // A different first-page header means a different amount of room on
            // page one. One box for the whole document is right for neither.
            const pages = layoutPages([para(lines(9))], PAGE, {
                widowOrphanControl: false,
                contentBox: (pageIndex) => ({
                    topPx: PAGE.marginTopPx,
                    // Page one holds three lines, the rest hold five.
                    bottomPx: PAGE.marginTopPx + (0 === pageIndex ? LINE * 3 : LINE * 5),
                }),
            });

            expect(shape(pages)).toEqual([3, 5, 1]);
        });

        it('starts each page at that page\'s own top', () => {
            const pages = layoutPages([para(lines(6))], PAGE, {
                widowOrphanControl: false,
                contentBox: (pageIndex) => ({
                    topPx: 0 === pageIndex ? 10 : 40,
                    bottomPx: PAGE.heightPx - 10,
                }),
            });

            expect(pages[0]!.lines[0]!.yPx).toBe(10);
            expect(pages[1]!.lines[0]!.yPx).toBe(40);
        });

        it('measures keep-together against the CURRENT page\'s height', () => {
            // A paragraph that fits a full page but not a shortened one must
            // still be placed rather than moved for ever.
            const pages = layoutPages([para(lines(4), { keepLinesTogether: true })], PAGE, {
                contentBox: () => ({ topPx: PAGE.marginTopPx, bottomPx: PAGE.marginTopPx + LINE * 2 }),
            });

            expect(pages.flatMap((page) => page.lines).length).toBe(4);
        });
    });

    describe('layoutSections', () => {
        describe('page fields', () => {
            const numbered = (text: string, field?: 'page' | 'numPages'): Paragraph => ({
                runs: [{ text, font: MONO, sizePx: 16, ...(undefined === field ? {} : { field }) }],
                style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
            });

            const footerText = (page: ReturnType<typeof layoutSections>[number]): string =>
                (page.footer?.lines ?? [])
                    .map((placed) => placed.line.pieces.map((piece) => piece.text).join(''))
                    .join('');

            it('gives NUMPAGES the DOCUMENT total, not the section count', () => {
                // Each section here is one page, so a section-local count would
                // print "1" on both. The whole reason the total is threaded
                // through rather than taken from `pages.length`.
                const section = (): Section => ({
                    blocks: [para('body')],
                    geometry: PAGE,
                    footerFor: () => [numbered('?', 'numPages')],
                });

                const pages = layoutSections([section(), section()]);

                expect(pages.map(footerText)).toEqual(['2', '2']);
            });

            it('does not lay the document out twice when nothing asks for the total', () => {
                // A PAGE field is answered exactly on the first pass, so it must
                // not trigger the second.
                let calls = 0;
                const section = (): Section => ({
                    blocks: [para('body')],
                    geometry: PAGE,
                    footerFor: () => {
                        calls++;

                        return [numbered('?', 'page')];
                    },
                });

                const pages = layoutSections([section(), section()]);

                expect(pages.map(footerText)).toEqual(['1', '2']);
                // Once per page to draw, once per page to check for NUMPAGES.
                expect(calls).toBe(4);
            });

            it('resolves a field before the line is measured, not after', () => {
                // "10" is wider than "1". Resolving after the break would leave
                // the text a digit wider than the line it was fitted to — so the
                // measured height and width have to come from the RESOLVED text.
                const wide = { ...PAGE, widthPx: CELL * 4 + 20 };
                const [page] = layoutPages([para('body')], wide, {
                    totalPages: 1000,
                    footerFor: () => [numbered('?', 'numPages')],
                });

                expect(footerText(page!)).toBe('1000');
                // Four characters exactly fill the four-cell column.
                expect(page!.footer?.lines.length).toBe(1);
                expect(page!.footer?.lines[0]!.line.widthPx).toBeCloseTo(CELL * 4, 6);
            });
        });

        describe('furniture across sections', () => {
            const PAPER: PageGeometry = { ...PAGE, headerDistancePx: 14 };

            /** A section filling `count` whole pages. */
            const run = (count: number, extra: Partial<Section> = {}): Section => ({
                blocks: [para(lines(count * 5))],
                geometry: PAPER,
                ...extra,
            });

            it('counts the page NUMBER through the whole document', () => {
                // Why the number is passed at all: section two's first page is
                // the document's fourth, and a header alternating by parity has
                // to know that. A section-relative index calls it page one.
                const asked: [number, number][] = [];
                const observe = (pageIndex: number, pageNumber: number): undefined => {
                    asked.push([pageIndex, pageNumber]);

                    return undefined;
                };

                const pages = layoutSections([
                    run(3, { headerFor: observe }),
                    run(1, { headerFor: observe }),
                ], { widowOrphanControl: false });

                expect(pages.map((page) => page.pageNumber)).toEqual([1, 2, 3, 4]);
                // The DISTINCT pairs are the contract. A multi-section document
                // is also probed once for `NUMPAGES` before it is drawn, so the
                // callback is asked twice per page; asserting the raw sequence
                // would pin that probe rather than the numbering.
                const distinct = [...new Set(asked.map((pair) => pair.join(':')))];
                expect(distinct).toEqual(['0:1', '1:2', '2:3', '0:4']);
            });

            it('gives the blank parity page NO furniture and still numbers it', () => {
                // Verified against LibreOffice: the page inserted to reach an
                // odd-page start prints entirely empty, while the pages either
                // side of it keep their headers.
                const pages = layoutSections([
                    run(1, { headerFor: () => [para('A')] }),
                    run(1, { startsOn: 'oddPage', headerFor: () => [para('B')] }),
                ], { widowOrphanControl: false });

                expect(pages.length).toBe(3);
                expect(pages[1]!.lines.length).toBe(0);
                expect(pages[1]!.header).toBeUndefined();
                expect(pages.map((page) => page.pageNumber)).toEqual([1, 2, 3]);
                expect(pages[2]!.header?.lines.length).toBe(1);
            });

            it('draws each section its OWN furniture', () => {
                const pages = layoutSections([
                    run(1, { headerFor: () => [para('AAA')] }),
                    run(1, { headerFor: () => [para('BBB')] }),
                ], { widowOrphanControl: false });

                const text = (index: number): string =>
                    pages[index]!.header!.lines[0]!.line.pieces
                        .map((piece) => piece.text).join('');

                expect([text(0), text(1)]).toEqual(['AAA', 'BBB']);
            });
        });

        const landscape: PageGeometry = { ...PAGE, widthPx: PAGE.heightPx, heightPx: PAGE.widthPx };

        it('starts each section on a new page and gives it its own paper', () => {
            const pages = layoutSections([
                { blocks: [para('a')], geometry: PAGE },
                { blocks: [para('b')], geometry: landscape },
            ]);

            expect(pages.length).toBe(2);
            expect(pages[0]!.geometry.widthPx).toBe(PAGE.widthPx);
            expect(pages[1]!.geometry.widthPx).toBe(landscape.widthPx);
            expect(pages.map((page) => page.sectionIndex)).toEqual([0, 1]);
        });

        it('numbers blocks across the WHOLE document, not within a section', () => {
            // A caller maps a placed line back to the block it came from without
            // knowing which section that was; section-relative indices would
            // point at the wrong paragraph in every section after the first.
            const pages = layoutSections([
                { blocks: [para('a'), para('b')], geometry: PAGE },
                { blocks: [para('c')], geometry: PAGE },
            ]);

            expect(pages[0]!.lines.map((line) => line.paragraphIndex)).toEqual([0, 1]);
            expect(pages[1]!.lines.map((line) => line.paragraphIndex)).toEqual([2]);
        });

        it('paginates a long section within its own paper', () => {
            const pages = layoutSections([
                { blocks: [para(lines(7))], geometry: PAGE },
                { blocks: [para('tail')], geometry: PAGE },
            ], { widowOrphanControl: false });

            expect(shape(pages)).toEqual([5, 2, 1]);
            expect(pages.map((page) => page.sectionIndex)).toEqual([0, 0, 1]);
        });

        it('offsets a TABLE row\'s block index too', () => {
            // Rows carry their own index, and a renderer maps it back the same
            // way a line's is mapped. Offsetting only the lines would point at
            // the wrong block for every table after the first section.
            const table = {
                rows: [{ cells: [{ paragraphs: [para('cell')], gridSpan: 1 }], isHeader: false, cantSplit: false }],
                columnWidthsPx: [CELL * 10],
                cellMarginLeftPx: 0,
                cellMarginRightPx: 0,
                cellMarginTopPx: 0,
                cellMarginBottomPx: 0,
                spaceBeforePx: 0,
                spaceAfterPx: 0,
                pageBreakBefore: false,
            };

            const pages = layoutSections([
                { blocks: [para('a'), para('b')], geometry: PAGE },
                { blocks: [table], geometry: PAGE },
            ]);

            expect(pages[1]!.rows[0]!.blockIndex).toBe(2);
        });

        it('honours a section\'s OWN writing area', () => {
            // Each section has its own headers, so each has its own amount of
            // room. One box for the document would be right for neither.
            const pages = layoutSections([
                { blocks: [para(lines(4))], geometry: PAGE },
                {
                    blocks: [para(lines(4))],
                    geometry: PAGE,
                    // Two lines to a page in the second section only.
                    contentBox: () => ({ topPx: PAGE.marginTopPx, bottomPx: PAGE.marginTopPx + LINE * 2 }),
                },
            ], { widowOrphanControl: false });

            expect(shape(pages)).toEqual([4, 2, 2]);
        });

        it('inserts a blank page so a section can start on an ODD one', () => {
            // How a chapter always opens on a right-hand page. The blank is
            // real: it is printed, and it counts.
            const pages = layoutSections([
                { blocks: [para('a'), para('b')], geometry: PAGE, startsOn: 'nextPage' },
                { blocks: [para('c')], geometry: PAGE, startsOn: 'oddPage' },
            ]);

            // Section one ends on page 1, so the next page would be number 2 —
            // even. A blank page takes that slot and the section opens on 3.
            expect(pages.length).toBe(3);
            expect(pages[1]!.lines).toEqual([]);
            expect(pages[1]!.rows).toEqual([]);
            expect(pages[2]!.sectionIndex).toBe(1);
        });

        it('inserts no blank page when the parity already suits', () => {
            const pages = layoutSections([
                { blocks: [para('a')], geometry: PAGE, startsOn: 'nextPage' },
                { blocks: [para('b')], geometry: PAGE, startsOn: 'evenPage' },
            ]);

            // One page used, so the next is number two — already even.
            expect(pages.length).toBe(2);
            expect(pages[1]!.lines.length).toBe(1);
        });

        it('never opens the document with a blank page', () => {
            // The first section has nothing to be even or odd relative to.
            const pages = layoutSections([
                { blocks: [para('a')], geometry: PAGE, startsOn: 'evenPage' },
            ]);

            expect(pages.length).toBe(1);
            expect(pages[0]!.lines.length).toBe(1);
        });

        it('gives the blank page to the section BEFORE the break', () => {
            // It is the tail of that run of pages, and it is printed on that
            // section's paper.
            const landscapeFirst: PageGeometry = { ...PAGE, widthPx: PAGE.heightPx, heightPx: PAGE.widthPx };
            const pages = layoutSections([
                { blocks: [para('a'), para('b')], geometry: landscapeFirst, startsOn: 'nextPage' },
                { blocks: [para('c')], geometry: PAGE, startsOn: 'oddPage' },
            ]);

            expect(pages[1]!.sectionIndex).toBe(0);
            expect(pages[1]!.geometry.widthPx).toBe(landscapeFirst.widthPx);
        });

        it('says which paper every page is on, for a renderer to draw', () => {
            const pages = layoutSections([{ blocks: [para('a')], geometry: landscape }]);

            expect(pages[0]!.geometry.heightPx).toBe(landscape.heightPx);
        });
    });
});

describe('keep with next', () => {
    const KEEP: Partial<ParagraphStyle> = { keepWithNext: true };

    /** A page one character wide, so a paragraph of N words is N lines tall. */
    const NARROW: PageGeometry = { ...PAGE, widthPx: CELL + 20 };

    // Trimmed: a wrapped line keeps the space that broke it, which is a fact
    // about line breaking and not about which page the line landed on.
    const textOf = (pages: readonly Page[]): string[][] => pages.map(
        (page) => page.lines.map(
            (placed) => placed.line.pieces.map((piece) => piece.text).join('').trim()));

    it('moves a paragraph down to stay with the one after it', () => {
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('d'), para('HEAD', KEEP), para('BODY')],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'd'], ['HEAD', 'BODY']]);
    });

    it('leaves it where it is when what follows fits beside it', () => {
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('HEAD', KEEP), para('BODY')],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'HEAD', 'BODY']]);
    });

    it('chains back through every paragraph that carries it', () => {
        // Only the LOWER heading is short of room; the one above it is kept
        // with that one in turn, so both have to travel.
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('H1', KEEP), para('H2', KEEP), para('B')],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c'], ['H1', 'H2', 'B']]);
    });

    it('needs only the FIRST line of what follows', () => {
        // What follows the heading is four lines long and cannot fit beside it.
        // Demanding all of it would drive the heading off a page that has room.
        //
        // Orphan control off, because it would move that first line down for
        // reasons of its own and hide whether the keep needed it to.
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('H', KEEP), para('w w w w')],
            NARROW,
            { widowOrphanControl: false },
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'H', 'w'], ['w', 'w', 'w']]);
    });

    it('moves a MULTI-LINE keep whole rather than splitting it', () => {
        // Measured: LibreOffice moved all five lines of a keepNext paragraph
        // rather than pushing down only the last one to meet what follows.
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('w w', KEEP), para('B')],
            NARROW,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c'], ['w', 'w', 'B']]);
    });

    it('measures the group by the room its lines will really take', () => {
        // A kept paragraph whose lines differ in height: 'small' at 16px is
        // 18.125 tall and 'BIG' at 32px is 36.25, so it needs 54.375 and its
        // group needs 74.375 with the follower's first line. There is 80 left
        // on this page, so it stays — and the whole group with it.
        //
        // Measuring the group at the paragraph's NOMINAL height instead — its
        // tallest run for every line, which is what a per-paragraph box
        // gives — makes the same group 92.5 and pushes it onto a page of
        // its own for room it never needed. No printed page decides this one:
        // the flow places lines at their own heights, so the keep has to ask
        // about the same heights or it is answering a different question.
        const tall: Paragraph = {
            runs: [
                { text: 'small ', font: MONO, sizePx: 16 },
                { text: 'BIG', font: MONO, sizePx: 32 },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, keepWithNext: true },
        };
        const pages = layoutPages([para('a'), tall, para('B')], PAGE);

        expect(textOf(pages)).toEqual([['a', 'small', 'BIG', 'B']]);
    });

    it('does nothing for a lone paragraph at the end of the document', () => {
        // There is nothing to keep it WITH, and breaking the page for that
        // would strand it alone on one of its own.
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('d'), para('LAST', KEEP)],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'd', 'LAST']]);
    });

    it('still keeps a CHAIN that reaches the end of the document', () => {
        // The last one has nothing to keep it with, but the one above it is
        // kept with the last — and that keep is as real as any other.
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('d'), para('H', KEEP), para('T', KEEP)],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'd'], ['H', 'T']]);
    });

    it('needs only the first line of the paragraph that ENDS such a chain', () => {
        // The heading is kept with a three-line paragraph that ends the
        // document. Standing in for the missing follower does not make that
        // paragraph indivisible: one line of it beside the heading is the
        // whole of what the keep asks for, and demanding all three would push
        // the heading off a page with room for it.
        const pages = layoutPages(
            [para('a'), para('b'), para('H', KEEP), para('w w w', KEEP)],
            NARROW,
            { widowOrphanControl: false },
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'H', 'w', 'w'], ['w']]);
    });

    it('stops the chain at a paragraph that opens its own page', () => {
        // The break is INSIDE the chain, not after it. The keep above it can
        // never be honoured, so counting the paragraphs beyond the break into
        // the group would push the heading off this page and leave a hole here
        // as well as the one the break itself makes.
        const pages = layoutPages(
            [
                para('a'), para('b'), para('c'), para('d'), para('H', KEEP),
                para('M', { ...KEEP, pageBreakBefore: true }), para('BODY'),
            ],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'd', 'H'], ['M', 'BODY']]);
    });

    it('opens no blank page for a group that starts one', () => {
        // Six bound paragraphs at the very top of the document do not fit on a
        // five-line page. There is nowhere better for them to go, and breaking
        // to find one would leave page one empty.
        const pages = layoutPages(
            [
                ...['K1', 'K2', 'K3', 'K4', 'K5', 'K6'].map((text) => para(text, KEEP)),
                para('END'),
            ],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['K1', 'K2', 'K3', 'K4', 'K5'], ['K6', 'END']]);
    });

    it('does nothing in front of a paragraph that opens its own page', () => {
        // The keep cannot be honoured whatever happens, so spending a page
        // break on it only leaves a hole above the one Word asked for.
        const pages = layoutPages(
            [
                para('a'), para('b'), para('c'), para('d'), para('HEAD', KEEP),
                para('BODY', { pageBreakBefore: true }),
            ],
            PAGE,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'd', 'HEAD'], ['BODY']]);
    });

    it('places a group too tall for any page rather than looping', () => {
        // Six bound paragraphs do not fit on a five-line page. LibreOffice
        // starts them on a fresh page and breaks inside the group, which is as
        // much of the keep as can survive — and it must not go on breaking.
        const pages = layoutPages(
            [
                para('a'), para('b'),
                ...['K1', 'K2', 'K3', 'K4', 'K5', 'K6'].map((text) => para(text, KEEP)),
                para('END'),
            ],
            PAGE,
        );

        expect(textOf(pages)).toEqual([
            ['a', 'b'],
            ['K1', 'K2', 'K3', 'K4', 'K5'],
            ['K6', 'END'],
        ]);
    });
});

describe('columns', () => {
    /** Two columns of four characters, with two characters between them. */
    const TWO: PageGeometry = {
        ...PAGE,
        widthPx: CELL * 10 + 20,
        columns: [
            { leftPx: 10, widthPx: CELL * 4 },
            { leftPx: 10 + CELL * 6, widthPx: CELL * 4 },
        ],
    };

    const textOf = (pages: readonly Page[]): string[][] => pages.map(
        (page) => page.lines.map(
            (placed) => placed.line.pieces.map((piece) => piece.text).join('').trim()));

    it('fills a column to the foot of the page before starting the next', () => {
        // Five lines to a column and two columns, so ten before a new page.
        const pages = layoutPages(
            Array.from({ length: 12 }, (_, index) => para(`L${index + 1}`)),
            TWO,
        );

        expect(textOf(pages)).toEqual([
            ['L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L10'],
            ['L11', 'L12'],
        ]);
    });

    it('puts each column at its own left edge', () => {
        const [page] = layoutPages(
            Array.from({ length: 7 }, (_, index) => para(`L${index + 1}`)),
            TWO,
        );
        const xs = page!.lines.map((placed) => placed.xPx);

        expect(xs.slice(0, 5)).toEqual([10, 10, 10, 10, 10]);
        expect(xs.slice(5)).toEqual([10 + CELL * 6, 10 + CELL * 6]);
    });

    it('breaks lines at the COLUMN width, not the writing width', () => {
        // Six characters fit the page and not the four-character column, so a
        // column that measured against the page would run into its neighbour.
        const [page] = layoutPages([para('aa bb')], TWO);

        expect(page!.lines.length).toBe(2);
    });

    it('does NOT balance the columns a section ends on', () => {
        // This was once read as a missing feature. LibreOffice ends a document
        // the same way: forty lines over two columns left the last page's
        // twelve all in the left-hand one, with the right empty.
        // Balancing happens at a CONTINUOUS section break, which is a
        // different thing and is reported unsupported by the reader.
        const pages = layoutPages(
            Array.from({ length: 12 }, (_, index) => para(`L${index + 1}`)),
            TWO,
        );
        const last = pages[pages.length - 1]!;

        expect(last.lines.every((placed) => placed.xPx === TWO.columns![0]!.leftPx)).toBe(true);
        expect(last.lines.length).toBe(2);
    });

    it('starts a new page back in the FIRST column', () => {
        // Twelve lines is two full columns and two over. A page that carried on
        // from the column the last one ended in would open every page in the
        // right-hand column and leave the left of it blank forever.
        const pages = layoutPages(
            Array.from({ length: 12 }, (_, index) => para(`L${index + 1}`)),
            TWO,
        );

        expect(pages[1]!.lines.map((placed) => placed.xPx)).toEqual([10, 10]);
    });

    it('places something too tall for a column in the empty one it is in', () => {
        // Nowhere better to go: moving on would leave the column blank and put
        // the paragraph in one exactly as short. Asking whether the PAGE is
        // empty would answer no here and waste a column every time.
        const tall = para('x', { lineHeightPx: 120 });
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('d'), para('e'), tall],
            TWO,
        );

        expect(pages.length).toBe(1);
        expect(pages[0]!.lines[5]!.xPx).toBe(10 + CELL * 6);
    });

    it('draws a table in the column it lands in', () => {
        const cell = { paragraphs: [para('t')], gridSpan: 1 };
        const pages = layoutPages(
            [
                para('a'), para('b'), para('c'), para('d'), para('e'),
                {
                    rows: [{ cells: [cell], isHeader: false, cantSplit: false }],
                    columnWidthsPx: [CELL * 4],
                    cellMarginLeftPx: 0,
                    cellMarginRightPx: 0,
                    cellMarginTopPx: 0,
                    cellMarginBottomPx: 0,
                    spaceBeforePx: 0,
                    spaceAfterPx: 0,
                    pageBreakBefore: false,
                },
            ],
            TWO,
        );

        expect(pages.length).toBe(1);
        expect(pages[0]!.rows[0]!.cells[0]!.xPx).toBe(10 + CELL * 6);
    });

    it('measures a header against the WRITING width, not a column', () => {
        // Furniture spans the page however many columns the body has. Six
        // characters fit the ten-character writing width and not the
        // four-character column, so a header measured against a column would
        // wrap where the document does not.
        const [page] = layoutPages([para('body')], TWO, {
            headerFor: () => [para('aa bb')],
        });

        expect(page!.header?.lines.length).toBe(1);
    });

    /** A wide first column and a narrow second one. */
    const UNEVEN: PageGeometry = {
        ...PAGE,
        widthPx: CELL * 30 + 20,
        columns: [
            { leftPx: 10, widthPx: CELL * 10 },
            { leftPx: 10 + CELL * 12, widthPx: CELL * 4 },
        ],
    };

    /**
     * A paragraph of SEVERAL runs, long enough to fill the wide column and
     * carry on into the narrow one — so the split falls inside a run that is
     * not the first, which is the only place the arithmetic can go wrong.
     */
    const spanning = (): Paragraph => ({
        runs: Array.from({ length: 6 }, () => ({
            text: 'aa bb cc dd ',
            font: MONO,
            sizePx: 16,
        })),
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
    });

    it('breaks what carries into a NARROWER column again at its width', () => {
        // A paragraph is broken at one column's width; the lines it carried
        // into the next used to be drawn at that width, off the right edge of
        // the narrower one and off the paper with it.
        const [page] = layoutPages([spanning()], UNEVEN, { widowOrphanControl: false });
        const second = UNEVEN.columns![1]!;

        expect(page!.lines.some((placed) => placed.xPx >= second.leftPx)).toBe(true);

        for (const placed of page!.lines) {
            if (placed.xPx >= second.leftPx) {
                expect(placed.line.widthPx).toBeLessThanOrEqual(second.widthPx);
            }
        }
    });

    it('loses none of the text it re-broke, and repeats none of it', () => {
        // Re-breaking the WHOLE paragraph would place its opening twice;
        // re-breaking nothing would drop everything after the column change.
        const paragraph = spanning();
        const words = paragraph.runs.map((run) => run.text).join('').trim();
        const printed = layoutPages([paragraph], UNEVEN, { widowOrphanControl: false })
            .flatMap((page) => page.lines)
            .flatMap((placed) => placed.line.pieces.map((piece) => piece.text))
            .join('')
            .replace(/\s+/gu, ' ')
            .trim();

        expect(printed).toBe(words);
    });

    it('leaves a paragraph alone where the columns are the SAME width', () => {
        // Re-breaking at a width it was already broken at is work for nothing,
        // and the lines it has are already right.
        const paragraph = spanning();
        const words = paragraph.runs.map((run) => run.text).join('').trim();
        const even = layoutPages([paragraph], TWO, { widowOrphanControl: false });
        const printed = even.flatMap((page) => page.lines)
            .flatMap((placed) => placed.line.pieces.map((piece) => piece.text))
            .join('')
            .replace(/\s+/gu, ' ')
            .trim();

        expect(printed).toBe(words);
    });

    it('does not draw the list marker twice over', () => {
        // The marker belongs to the paragraph's FIRST line. What carries into
        // the next column is a fresh flow, and its first line is not that one.
        const item: Paragraph = {
            ...spanning(),
            marker: { run: { text: '1.', font: MONO, sizePx: 16 }, offsetPx: CELL },
        };
        const [page] = layoutPages([item], UNEVEN, { widowOrphanControl: false });
        const markers = page!.lines.filter((placed) => undefined !== placed.marker);

        expect(markers.length).toBe(1);
        expect(markers[0]).toBe(page!.lines[0]);
    });

    it('sends a page break to the next PAGE, not the next column', () => {
        // A column break and a page break are different instructions, and a
        // page break that only moved across would leave the rest of the page
        // blank AND fail to start a new one.
        const pages = layoutPages(
            [para('a'), para('b', { pageBreakBefore: true })],
            TWO,
        );

        expect(textOf(pages)).toEqual([['a'], ['b']]);
    });

    it('moves a keep group to the next column before the next page', () => {
        const pages = layoutPages(
            [
                para('a'), para('b'), para('c'), para('d'),
                para('H', { keepWithNext: true }), para('B'),
            ],
            TWO,
        );

        expect(textOf(pages)).toEqual([['a', 'b', 'c', 'd', 'H', 'B']]);
    });

    it('carries ONE paragraph on into the next column', () => {
        // Seven three-letter words in a four-character column is seven lines,
        // and a column holds five. The two that overflow are drawn where the
        // next column is — a paragraph measured in one column and drawn at its
        // left edge throughout would print the tail of it over the first
        // column's own text.
        const [page] = layoutPages([para('aaa aaa aaa aaa aaa aaa aaa')], TWO);

        expect(page!.lines.length).toBe(7);
        expect(page!.lines.map((placed) => placed.xPx))
            .toEqual([10, 10, 10, 10, 10, 10 + CELL * 6, 10 + CELL * 6]);
        // ...and the sixth line starts at the TOP of the page again.
        expect(page!.lines[5]!.yPx).toBe(page!.lines[0]!.yPx);
    });
});

describe('character spacing', () => {
    const tracked = (letterSpacingPx?: number): Paragraph => ({
        runs: [{
            text: 'aaa bbb',
            font: MONO,
            sizePx: 16,
            ...(undefined === letterSpacingPx ? {} : { letterSpacingPx }),
        }],
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
    });

    it('breaks the line where the TRACKED width says to', () => {
        // The point of reading it at all. Seven characters fit the
        // ten-character page bare; tracked by a whole cell each they come to
        // fourteen and do not, and a run measured without its spacing would
        // run past the margin.
        expect(layoutPages([tracked()], PAGE)[0]!.lines.length).toBe(1);
        expect(layoutPages([tracked(CELL)], PAGE)[0]!.lines.length).toBe(2);
    });

    it('tracks the trailing spaces too, and leaves them out of the CONTENT', () => {
        // A line's content width is what alignment ranges on, and its trailing
        // spaces are not part of it. Measuring them bare would leave the
        // tracking on those spaces inside the content and range every centred
        // line a little left.
        const trailing = (letterSpacingPx?: number): Paragraph => ({
            runs: [{
                text: 'aa  ',
                font: MONO,
                sizePx: 16,
                ...(undefined === letterSpacingPx ? {} : { letterSpacingPx }),
            }],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
        });

        const [wide] = layoutPages([trailing(CELL)], { ...PAGE, widthPx: CELL * 40 });
        const line = wide!.lines[0]!.line;

        // Four characters tracked by a cell each, less the last two — which
        // are spaces, and are tracked too. Measuring those bare would leave
        // their tracking inside the width and range every centred line left.
        expect(line.widthPx).toBeCloseTo(CELL * 4, 6);
    });

    it('ranges a DECIMAL tab on the tracked width of what precedes the point', () => {
        // The prefix decides where the number hangs. Measured bare, a tracked
        // one lands short by its own tracking and the column stops lining up.
        const stop = { positionPx: CELL * 8, align: 'decimal' as const };
        const numbered = (letterSpacingPx?: number): Paragraph => ({
            runs: [{
                text: '\t12.5',
                font: MONO,
                sizePx: 16,
                ...(undefined === letterSpacingPx ? {} : { letterSpacingPx }),
            }],
            style: {
                ...DEFAULT_PARAGRAPH_STYLE,
                lineHeightPx: LINE,
                tabStops: [stop],
            },
        });

        const bare = layoutPages([numbered()], { ...PAGE, widthPx: CELL * 40 })[0]!;
        const wide = layoutPages([numbered(CELL)], { ...PAGE, widthPx: CELL * 40 })[0]!;
        const tabOf = (page: typeof bare): number => page.lines[0]!.line.pieces[0]!.widthPx;

        // '12' tracked by a cell each is two cells wider, so the tab that puts
        // the point on the stop is two cells SHORTER.
        expect(tabOf(bare) - tabOf(wide)).toBeCloseTo(CELL * 2, 6);
    });

    it('adds the space after EVERY character, the last one included', () => {
        // Word's own model, and what LibreOffice printed at half a point, one,
        // two and five. Seven characters tracked by one cell each is seven
        // cells wider than the same seven bare.
        const [wide] = layoutPages([tracked(CELL)], { ...PAGE, widthPx: CELL * 40 });
        const [bare] = layoutPages([tracked()], { ...PAGE, widthPx: CELL * 40 });

        expect(wide!.lines[0]!.line.widthPx - bare!.lines[0]!.line.widthPx)
            .toBeCloseTo(CELL * 7, 6);
    });
});

describe('footnotes', () => {
    const NOTE = (text: string): Paragraph[] => [para(text)];

    /** A paragraph whose one run is a reference mark for `id`. */
    const marked = (text: string, id: number): Paragraph => ({
        runs: [
            { text, font: MONO, sizePx: 16 },
            { text: String(id), font: MONO, sizePx: 16, footnoteId: id },
        ],
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
    });

    const withNotes = (blocks: Paragraph[], notes: Map<number, Paragraph[]>) =>
        layoutPages(blocks, PAGE, { footnotes: notes, widowOrphanControl: false });

    it('takes no room at all while the page carries no note', () => {
        const notes = new Map([[1, NOTE('n')]]);

        expect(withNotes([para('a'), para('b')], notes)[0]!.footnotes).toBeUndefined();
        expect(shape(withNotes(Array.from({ length: 5 }, () => para('x')), notes)))
            .toEqual([5]);
    });

    it('gives the note a line AND the rule a line of its own', () => {
        // Five lines fit a bare page. One note costs the note's line and the
        // line its rule stands on, so three body lines are left — measured
        // against LibreOffice, where fourteen became twelve.
        const notes = new Map([[1, NOTE('n')]]);
        const pages = withNotes(
            [para('a'), para('b'), para('c'), marked('d', 1), para('e')],
            notes,
        );

        expect(shape(pages)).toEqual([3, 2]);
    });

    it('moves the referencing line, not the line after it', () => {
        // The line that brings the note down is the one that has to fit beside
        // it. A page that reserved the room only once the line was placed would
        // print the note over the text.
        const notes = new Map([[1, NOTE('n')]]);
        const [page] = withNotes(
            [para('a'), para('b'), para('c'), para('d'), marked('e', 1)],
            notes,
        );

        expect(page!.lines.map((placed) => placed.line.pieces[0]!.text))
            .toEqual(['a', 'b', 'c', 'd']);
        expect(page!.footnotes).toBeUndefined();
    });

    it('stacks the notes against the FOOT of the page', () => {
        // Not under the body: measured against LibreOffice, the last note's
        // line sits on the bottom margin however short the page's text is.
        const notes = new Map([[1, NOTE('one')], [2, NOTE('two')]]);
        const [page] = withNotes([marked('a', 1), marked('b', 2)], notes);
        const foot = PAGE.heightPx - PAGE.marginBottomPx;
        const placed = page!.footnotes!;

        expect(placed.lines.length).toBe(2);
        expect(placed.lines[1]!.yPx + placed.lines[1]!.heightPx).toBe(foot);
        expect(placed.lines[0]!.yPx).toBe(foot - LINE * 2);
    });

    it('draws the rule two inches wide, above the notes', () => {
        const notes = new Map([[1, NOTE('one')]]);
        const [page] = withNotes([marked('a', 1)], notes);
        const foot = PAGE.heightPx - PAGE.marginBottomPx;
        const placed = page!.footnotes!;

        expect(placed.separatorWidthPx).toBe(FOOTNOTE_RULE_WIDTH_PX);
        expect(placed.separatorLeftPx).toBe(PAGE.marginLeftPx);
        expect(placed.separatorYPx).toBeCloseTo(foot - LINE - FOOTNOTE_RULE_GAP_PX, 6);
    });

    it('keeps each note on the page its own reference landed on', () => {
        const notes = new Map([[1, NOTE('one')], [2, NOTE('two')]]);
        const pages = withNotes(
            [marked('a', 1), para('b'), para('c'), para('d'), para('e'), marked('f', 2)],
            notes,
        );

        expect(pages.length).toBe(2);
        expect(pages[0]!.footnotes!.lines[0]!.line.pieces[0]!.text).toBe('one');
        expect(pages[1]!.footnotes!.lines[0]!.line.pieces[0]!.text).toBe('two');
    });

    it('counts a note once however many marks point at it', () => {
        const notes = new Map([[1, NOTE('one')]]);
        const [page] = withNotes([marked('a', 1), marked('b', 1)], notes);

        expect(page!.footnotes!.lines.length).toBe(1);
    });

    /** A table whose one cell holds `blocks`. */
    const celled = (...blocks: Paragraph[]) => ({
        rows: [{
            cells: [{ paragraphs: blocks, gridSpan: 1 }],
            isHeader: false,
            cantSplit: false,
        }],
        columnWidthsPx: [CELL * 10],
        cellMarginLeftPx: 0,
        cellMarginRightPx: 0,
        cellMarginTopPx: 0,
        cellMarginBottomPx: 0,
        spaceBeforePx: 0,
        spaceAfterPx: 0,
        pageBreakBefore: false,
    });

    it('carries a row across the page counting each cell’s OWN top margin', () => {
        // The split has to know where the text starts inside the cell, and
        // that is the cell's margin and not the table's once `w:tcMar` has
        // said otherwise. With the two confused, the upper half comes back a
        // margin short and the lower half draws its first line over its own
        // padding.
        //
        // Two cells, one padded and one not, so a single row-wide number
        // cannot satisfy both.
        const lines = (count: number): Paragraph =>
            para(Array.from({ length: count }, () => 'xxxxxxxxxx').join(' '));
        const padded = {
            rows: [{
                cells: [
                    { paragraphs: [lines(6)], gridSpan: 1, margins: { topPx: LINE } },
                    { paragraphs: [lines(6)], gridSpan: 1 },
                ],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [CELL * 10, CELL * 10],
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };

        // A line ahead of it, because a row is only ever cut to fill a page
        // that already has something on it — one alone is placed whole.
        const pages = layoutPages([para('ahead'), padded], PAGE);

        expect(pages.length).toBeGreaterThan(1);

        const upper = pages[0]!.rows[0]!;
        const lower = pages[1]!.rows[0]!;

        // The padded cell's first line starts a whole margin below the row,
        // and the unpadded one's starts at the top of it.
        expect(upper.cells[0]!.lines[0]!.yPx - upper.yPx).toBe(LINE);
        expect(upper.cells[1]!.lines[0]!.yPx - upper.yPx).toBe(0);

        // The padded cell fits one line fewer in the same room, and what is
        // left of it lands below its own padding on the next page — not at the
        // top of the row, and not overlapping what stayed behind.
        expect(upper.cells[0]!.lines.length).toBe(upper.cells[1]!.lines.length - 1);
        expect(lower.cells[0]!.lines[0]!.yPx - lower.yPx).toBe(LINE);

        // Nothing is lost or drawn twice: six lines went in and six come out.
        expect(upper.cells[0]!.lines.length + lower.cells[0]!.lines.length).toBe(6);
        expect(upper.cells[1]!.lines.length + lower.cells[1]!.lines.length).toBe(6);
    });

    it('measures the upper half of a split row INCLUDING the margin above it', () => {
        // The test above cannot see this: with an unpadded cell beside it, the
        // unpadded one is the taller half and decides the cut on its own. One
        // padded cell alone is what makes the margin part of the answer — and
        // getting it wrong slides every carried line down by a margin, because
        // the padding is added again when the row is drawn.
        const alone = {
            rows: [{
                cells: [{
                    paragraphs: [para(Array.from({ length: 6 }, () => 'xxxxxxxxxx').join(' '))],
                    gridSpan: 1,
                    margins: { topPx: LINE },
                }],
                isHeader: false,
                cantSplit: false,
            }],
            columnWidthsPx: [CELL * 10],
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };

        const pages = layoutPages([para('ahead'), alone], PAGE);
        const upper = pages[0]!.rows[0]!;
        const lower = pages[1]!.rows[0]!;

        // Three lines fit in the 80px left, once the 20px margin is taken off.
        expect(upper.cells[0]!.lines.length).toBe(3);
        expect(upper.heightPx).toBe(LINE * 3 + LINE);
        // And the carried text lands one margin down its own row, not two.
        expect(lower.cells[0]!.lines[0]!.yPx - lower.yPx).toBe(LINE);
    });

    it('keeps the note of a reference made inside a TABLE CELL', () => {
        // A cell is a little page of its own, and a reference in one used to
        // draw its mark and lose the note it pointed at.
        const [page] = layoutPages([celled(marked('a', 1))], PAGE, {
            footnotes: new Map([[1, NOTE('celled')]]),
        });

        expect(page!.footnotes?.lines.map(
            (placed) => placed.line.pieces.map((piece) => piece.text).join('')))
            .toEqual(['celled']);
    });

    it('finds one nested as deep as the document puts it', () => {
        const inner = celled(marked('deep', 1));
        const outer = {
            ...celled(),
            rows: [{
                cells: [{ paragraphs: [inner], gridSpan: 1 }],
                isHeader: false,
                cantSplit: false,
            }],
        };
        const [page] = layoutPages([outer], PAGE, {
            footnotes: new Map([[1, NOTE('deeper')]]),
        });

        expect(page!.footnotes?.lines.length).toBe(1);
    });

    it('moves the ROW when its note leaves no room for it', () => {
        // The same rule a paragraph gets: the row that brings the note down
        // has to fit beside it, or the page it would have shared is the wrong
        // one for both.
        const notes = new Map([[1, NOTE('n')]]);
        const pages = layoutPages(
            [para('a'), para('b'), para('c'), para('d'), celled(marked('e', 1))],
            PAGE,
            { footnotes: notes, widowOrphanControl: false },
        );

        expect(pages[0]!.rows.length).toBe(0);
        expect(pages[0]!.footnotes).toBeUndefined();
        expect(pages[1]!.rows.length).toBe(1);
        expect(pages[1]!.footnotes?.lines.length).toBe(1);
    });

    it('owes one note once however many ROWS point at it', () => {
        // Two rows of a table citing the same note. Owing it per row shrinks
        // the page by a note that is only going to be printed once, and the
        // third row falls off for nothing.
        const notes = new Map([[1, NOTE('n')]]);
        const rowWith = (block: Paragraph) => ({
            cells: [{ paragraphs: [block], gridSpan: 1 }],
            isHeader: false,
            cantSplit: false,
        });
        const pages = layoutPages(
            [{
                ...celled(),
                rows: [rowWith(marked('a', 1)), rowWith(marked('b', 1)), rowWith(para('c'))],
            }],
            PAGE,
            { footnotes: notes, widowOrphanControl: false },
        );

        expect(pages.length).toBe(1);
        expect(pages[0]!.rows.length).toBe(3);
        expect(pages[0]!.footnotes?.lines.length).toBe(1);
    });

    it('does not charge a row for a note the page already owes', () => {
        // The paragraph above it brought the note down. Counting it again
        // while deciding whether the row fits reserves two notes' worth for
        // one note.
        const notes = new Map([[1, NOTE('n')]]);
        const pages = layoutPages(
            [marked('a', 1), para('b'), celled(marked('c', 1))],
            PAGE,
            { footnotes: notes, widowOrphanControl: false },
        );

        expect(pages.length).toBe(1);
        expect(pages[0]!.rows.length).toBe(1);
    });

    it('cuts a row at the room its OWN note leaves', () => {
        // The row is split, and the note it carries is part of what it has to
        // fit beside: cutting at the foot of the paper would leave lines
        // sitting where the note is about to be printed.
        const notes = new Map([[1, NOTE('n')]]);
        const tall = {
            ...celled(),
            rows: [{
                cells: [{
                    paragraphs: [marked('a', 1), para('b'), para('c'), para('d')],
                    gridSpan: 1,
                }],
                isHeader: false,
                cantSplit: false,
            }],
        };
        const pages = layoutPages([para('x'), tall], PAGE, {
            footnotes: notes,
            widowOrphanControl: false,
        });
        const first = pages[0]!.rows[0]!;

        // One line of text, then the row: two of its four lines fit above the
        // note and the rule, and the other two are carried over.
        expect(first.cells[0]!.lines.length).toBe(2);
        expect(first.yPx + first.heightPx)
            .toBeLessThanOrEqual(pages[0]!.footnotes!.separatorYPx);
        expect(pages[1]!.rows[0]!.cells[0]!.lines.length).toBe(2);
    });

    it('does not owe the note of a repeated header twice', () => {
        // A repeated header is the same row drawn again. Owing its note on
        // every continuation would shrink each page for a note already at the
        // foot of the first.
        const notes = new Map([[1, NOTE('n')]]);
        const header = {
            cells: [{ paragraphs: [marked('H', 1)], gridSpan: 1 }],
            isHeader: true,
            cantSplit: false,
        };
        const body = {
            cells: [{ paragraphs: [para('x')], gridSpan: 1 }],
            isHeader: false,
            cantSplit: false,
        };
        const pages = layoutPages(
            [{ ...celled(), rows: [header, ...Array.from({ length: 8 }, () => body)] }],
            PAGE,
            { footnotes: notes, widowOrphanControl: false },
        );

        expect(pages.length).toBeGreaterThan(1);
        expect(pages[0]!.footnotes?.lines.length).toBe(1);
        expect(pages[1]!.footnotes).toBeUndefined();
    });

    /** A note of `count` lines, each its own paragraph. */
    const bigNote = (count: number): Paragraph[] =>
        Array.from({ length: count }, (_, index) => para(`n${index + 1}`));

    it('CONTINUES a note too tall for the page rather than drawing it over the text', () => {
        // Ten lines of note on a five-line page. Drawing it whole put it over
        // the body and off the top of the paper; every line of it is printed
        // now, and none of them where the text is.
        const pages = layoutPages(
            [para('a'), marked('b', 1), para('c')],
            PAGE,
            { footnotes: new Map([[1, bigNote(10)]]), widowOrphanControl: false },
        );
        const noteLines = pages.reduce(
            (sum, page) => sum + (page.footnotes?.lines.length ?? 0), 0);

        expect(noteLines).toBe(10);
        for (const page of pages) {
            const lowest = Math.max(
                PAGE.marginTopPx,
                ...page.lines.map((placed) => placed.yPx + placed.heightPx),
            );

            expect(page.footnotes?.separatorYPx ?? Infinity).toBeGreaterThanOrEqual(lowest);
        }
    });

    it('stops the body at the line that brought such a note', () => {
        // Measured against LibreOffice: the referencing line stayed where it
        // was and the note took everything below it, so what followed moved on.
        const pages = layoutPages(
            [para('a'), marked('b', 1), para('c')],
            PAGE,
            { footnotes: new Map([[1, bigNote(10)]]), widowOrphanControl: false },
        );
        const textOf = (page: Page): string[] => page.lines.map(
            (placed) => placed.line.pieces.map((piece) => piece.text).join(''));

        expect(textOf(pages[0]!)).toEqual(['a', 'b1']);
        expect(textOf(pages[1]!)).toEqual(['c']);
    });

    it('adds pages of nothing but note when the document runs out', () => {
        // The last page's own notes still have to be printed. Without somewhere
        // to carry them the tail of a long note is simply lost.
        const pages = layoutPages(
            [marked('a', 1)],
            PAGE,
            { footnotes: new Map([[1, bigNote(12)]]), widowOrphanControl: false },
        );
        const noteOnly = pages.filter((page) => 0 === page.lines.length);

        expect(pages.length).toBeGreaterThan(1);
        expect(noteOnly.length).toBeGreaterThan(0);
        expect(pages.reduce((sum, page) => sum + (page.footnotes?.lines.length ?? 0), 0))
            .toBe(12);
    });

    it('prints what it carried BEFORE the notes of the page it carried it to', () => {
        // A continued note is still the earlier note: putting the new one first
        // would print the two out of order.
        const pages = layoutPages(
            [marked('a', 1), para('b'), para('c'), para('d'), marked('e', 2)],
            PAGE,
            {
                footnotes: new Map([[1, bigNote(8)], [2, NOTE('second')]]),
                widowOrphanControl: false,
            },
        );
        const printed = pages.flatMap((page) => page.footnotes?.lines ?? [])
            .map((placed) => placed.line.pieces.map((piece) => piece.text).join(''));

        expect(printed[printed.length - 1]).toBe('second');
        expect(printed.slice(0, 8)).toEqual(bigNote(8).map(
            (block) => block.runs[0]!.text));
    });

    it('keeps room on the page it carried a note ON TO', () => {
        // The continuation is owed before anything of that page's own. A page
        // that filled itself with text first would push the tail of the note
        // along in front of it and never print it.
        const pages = layoutPages(
            [marked('a', 1), ...Array.from({ length: 8 }, (_, index) => para(`b${index}`))],
            PAGE,
            { footnotes: new Map([[1, bigNote(7)]]), widowOrphanControl: false },
        );

        // The second page is still paying for the first page's note, so it
        // holds less text than a page with no note to carry.
        // Page two prints the rest of the note AND leaves room for it: the
        // text it holds plus the note it carried comes to no more than a page.
        const carried = pages[1]!.footnotes!.lines.length;

        expect(carried).toBeGreaterThan(0);
        expect(pages[1]!.lines.length + carried).toBeLessThanOrEqual(4);
    });

    it('stacks two notes on one page in order, one below the other', () => {
        // Each note is laid out from its own origin, so they all begin at
        // nought: printed without being re-based they would sit on top of one
        // another at the foot of the page.
        // The FIRST note is two lines, so the second note's own line begins at
        // nought while the line above it does not — which is the only shape
        // that can tell a re-based stack from one left where it lay.
        const notes = new Map([[1, bigNote(2)], [2, NOTE('two')]]);
        const [page] = layoutPages(
            [marked('a', 1), marked('b', 2)],
            { ...PAGE, heightPx: LINE * 10 + 20 },
            { footnotes: notes, widowOrphanControl: false },
        );
        const placed = page!.footnotes!.lines;

        expect(placed.map((line) => line.line.pieces[0]!.text)).toEqual(['n1', 'n2', 'two']);
        expect(placed[1]!.yPx).toBe(placed[0]!.yPx + placed[0]!.heightPx);
        expect(placed[2]!.yPx).toBe(placed[1]!.yPx + placed[1]!.heightPx);
    });

    it('rules a CONTINUED note the same two inches as any other', () => {
        // Word has a `continuationSeparator` for this, and its absence here
        // read as though something were missing. LibreOffice drew the same
        // 144pt rule on the continuation page as on the first — measured off
        // both streams of the same PDF — so there is nothing else to draw.
        const pages = layoutPages(
            [para('a'), marked('b', 1), para('c')],
            PAGE,
            { footnotes: new Map([[1, bigNote(10)]]), widowOrphanControl: false },
        );
        const ruled = pages.filter((page) => undefined !== page.footnotes);

        expect(ruled.length).toBeGreaterThan(1);
        for (const page of ruled) {
            expect(page.footnotes!.separatorWidthPx).toBe(FOOTNOTE_RULE_WIDTH_PX);
            expect(page.footnotes!.separatorLeftPx).toBe(PAGE.marginLeftPx);
        }
    });

    it('counts a note once when ONE line carries two marks for it', () => {
        // A note referenced twice in the same sentence is still one note, and
        // reserving for it twice would push a line off the page for nothing.
        const twice: Paragraph = {
            runs: [
                { text: 'a', font: MONO, sizePx: 16 },
                { text: '1', font: MONO, sizePx: 16, footnoteId: 1 },
                { text: 'b', font: MONO, sizePx: 16 },
                { text: '1', font: MONO, sizePx: 16, footnoteId: 1 },
            ],
            style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
        };
        // Three lines fit beside one note. Reserving twice over would leave
        // room for two and push the referencing line off the page.
        const pages = layoutPages([para('x'), para('y'), twice], PAGE, {
            footnotes: new Map([[1, NOTE('one')]]),
            widowOrphanControl: false,
        });

        expect(pages.length).toBe(1);
        expect(pages[0]!.lines.length).toBe(3);
        expect(pages[0]!.footnotes!.lines.length).toBe(1);
    });

    it('keeps a TABLE off the notes as well as text', () => {
        // Everything that asks where the page ends has to get the shortened
        // answer, or a table placed after a footnoted line prints over it.
        const notes = new Map([[1, NOTE('one')]]);
        const row = {
            cells: [{ paragraphs: [para('t')], gridSpan: 1 }],
            isHeader: false,
            cantSplit: false,
        };
        const table = {
            rows: [row, row],
            columnWidthsPx: [CELL * 4],
            cellMarginLeftPx: 0,
            cellMarginRightPx: 0,
            cellMarginTopPx: 0,
            cellMarginBottomPx: 0,
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };
        const pages = layoutPages([marked('a', 1), para('b'), table], PAGE, {
            footnotes: notes,
            widowOrphanControl: false,
        });
        const foot = pages[0]!.footnotes!.separatorYPx;

        for (const placed of pages[0]!.rows) {
            expect(placed.yPx + placed.heightPx).toBeLessThanOrEqual(foot);
        }
    });
});

describe('a section that restarts its page numbering', () => {
    const sectionOf = (blocks: Paragraph[], firstPageNumber?: number): Section => ({
        blocks,
        geometry: PAGE,
        startsOn: 'nextPage',
        ...(undefined === firstPageNumber ? {} : { firstPageNumber }),
    });

    it('numbers its pages from the number it was given', () => {
        // Not from how many pages came before it, which is the one thing the
        // document said it is not: the second section prints 1 again.
        const pages = layoutSections([
            sectionOf([para('a'), para('b')]),
            sectionOf([para('c')], 1),
        ]);

        expect(pages.map((page) => page.pageNumber)).toEqual([1, 1]);
    });

    it('carries on counting where the last section left off without one', () => {
        const pages = layoutSections([
            sectionOf([para('a')]),
            sectionOf([para('b')]),
        ]);

        expect(pages.map((page) => page.pageNumber)).toEqual([1, 2]);
    });

    it('keeps counting from the restart, not just at it', () => {
        const pages = layoutSections([
            sectionOf([para('a')]),
            sectionOf([para(lines(6))], 5),
        ]);

        expect(pages.map((page) => page.pageNumber)).toEqual([1, 5, 6]);
    });

    describe('paragraph borders', () => {
        // Every number here was read off a PDF LibreOffice printed from a
        // one-point (`w:sz="8"`) box on A4 with a one-centimetre margin. The
        // text sat at x=28.45 and the side rules at 27.9 and 567.45, against a
        // content edge of 28.35 and 566.93 — half a width outside, both sides.
        const side = (widthPx: number, spacePx = 0): BorderSide => ({
            widthPx,
            style: 'single' as BorderStyle,
            colorHex: '#000000',
            ...(0 === spacePx ? {} : { spacePx }),
        });

        const boxOf = (widthPx: number, spacePx = 0): BoxBorders => ({
            top: side(widthPx, spacePx),
            left: side(widthPx, spacePx),
            bottom: side(widthPx, spacePx),
            right: side(widthPx, spacePx),
        });

        const CONTENT_LEFT = 10;
        const CONTENT_RIGHT = PAGE.widthPx - 10;

        /** The step between two ordinary paragraphs, to measure a rule against. */
        const plainStep = (): number => {
            const page = layoutPages([para('a'), para('b')], PAGE)[0]!;

            return page.lines[1]!.yPx - page.lines[0]!.yPx;
        };

        it('puts each rule half its width outside the line box', () => {
            const [page] = layoutPages([para('xx', { borders: boxOf(2) })], PAGE);
            const [box] = page!.paragraphBorders;
            const [line] = page!.lines;

            expect(page!.paragraphBorders).toHaveLength(1);
            expect(box!.leftPx).toBe(CONTENT_LEFT - 1);
            expect(box!.rightPx).toBe(CONTENT_RIGHT + 1);
            expect(box!.topPx).toBe(line!.yPx - 1);
            expect(box!.bottomPx).toBe(line!.yPx + line!.heightPx + 1);
            expect(box!.opensHere).toBe(true);
            expect(box!.closesHere).toBe(true);
        });

        it('pushes every edge outward by `w:space`, the text staying put', () => {
            // LibreOffice moved the sides of a six-point-spaced box from 27.9
            // and 567.45 out to 21.9 and 573.45 — six points each — and left
            // the text at x=28.45 where it was. The box grows, not the column.
            const [page] = layoutPages([para('xx', { borders: boxOf(2, 6) })], PAGE);
            const [box] = page!.paragraphBorders;
            const [line] = page!.lines;

            expect(line!.xPx).toBe(CONTENT_LEFT);
            expect(box!.leftPx).toBe(CONTENT_LEFT - 6 - 1);
            expect(box!.rightPx).toBe(CONTENT_RIGHT + 6 + 1);
            expect(box!.topPx).toBe(line!.yPx - 6 - 1);
            expect(box!.bottomPx).toBe(line!.yPx + line!.heightPx + 6 + 1);
        });

        it('takes the WHOLE width of a rule out of the flow, not half of it', () => {
            // The step from an ordinary paragraph to a bordered one measured
            // 12.5pt where an unbordered pair stepped 11.5 — the rule's whole
            // width, because it is the OUTER edge the text above has to clear
            // — and 18.5 once six points of space were added.
            const step = plainStep();
            const bordered = layoutPages(
                [para('a'), para('b', { borders: boxOf(2) })], PAGE)[0]!;
            const spaced = layoutPages(
                [para('a'), para('b', { borders: boxOf(2, 6) })], PAGE)[0]!;

            expect(bordered.lines[1]!.yPx - bordered.lines[0]!.yPx).toBe(step + 2);
            expect(spaced.lines[1]!.yPx - spaced.lines[0]!.yPx).toBe(step + 6 + 2);
        });

        it('gives the bottom rule its room before the paragraph below', () => {
            const page = layoutPages(
                [para('a', { borders: boxOf(2, 6) }), para('b')], PAGE)[0]!;

            expect(page.lines[1]!.yPx - page.lines[0]!.yPx).toBe(plainStep() + 6 + 2);
        });

        it('draws ONE box round adjacent paragraphs carrying the same border', () => {
            // The find that decided the shape of all this: LibreOffice printed
            // two identically bordered paragraphs as a single outline, with the
            // side rules meeting and NO rule between them — so the pair stepped
            // 11.5pt, exactly as an unbordered pair does.
            const borders = boxOf(2);
            const page = layoutPages(
                [para('a', { borders }), para('b', { borders })], PAGE)[0]!;

            expect(page.paragraphBorders).toHaveLength(1);
            expect(page.lines[1]!.yPx - page.lines[0]!.yPx).toBe(plainStep());

            const [box] = page.paragraphBorders;
            expect(box!.topPx).toBe(page.lines[0]!.yPx - 1);
            expect(box!.bottomPx).toBe(page.lines[1]!.yPx + page.lines[1]!.heightPx + 1);
        });

        it('compares borders by value, so two paragraphs need not share a style', () => {
            // The same border stated twice is two objects. LibreOffice merged
            // them all the same, so identity is the wrong test.
            const page = layoutPages(
                [para('a', { borders: boxOf(2) }), para('b', { borders: boxOf(2) })],
                PAGE)[0]!;

            expect(page.paragraphBorders).toHaveLength(1);
        });

        it('starts a new box where the border differs', () => {
            const page = layoutPages(
                [para('a', { borders: boxOf(2) }), para('b', { borders: boxOf(4) })],
                PAGE)[0]!;

            expect(page.paragraphBorders).toHaveLength(2);
            expect(page.paragraphBorders.map((box) => box.closesHere)).toEqual([true, true]);
        });

        it('ends the box at a paragraph with no border of its own', () => {
            const borders = boxOf(2);
            const page = layoutPages(
                [para('a', { borders }), para('b'), para('c', { borders })], PAGE)[0]!;

            expect(page.paragraphBorders).toHaveLength(2);
        });

        it('closes a split box on BOTH pages, complete either side', () => {
            // Measured, where this used to be assumed: a bordered paragraph
            // split over two pages drew FOUR rules on each of them — a whole
            // box either side, not one outline left open at the foot and
            // picked up at the head. A table ROW splits the other way, and the
            // analogy with one is what gets this wrong.
            const pages = layoutPages(
                [para(lines(6), { borders: boxOf(2) })], PAGE,
                { widowOrphanControl: false });

            expect(pages.map((page) => page.paragraphBorders.length)).toEqual([1, 1]);

            const [first] = pages[0]!.paragraphBorders;
            const [second] = pages[1]!.paragraphBorders;

            expect([first!.opensHere, first!.closesHere]).toEqual([true, true]);
            expect([second!.opensHere, second!.closesHere]).toEqual([true, true]);
        });

        it('measures the carried half against the page it lands on, not the one it left', () => {
            // The half overleaf starts where the text starts THERE. Carrying
            // the top down from the page before would draw it off the paper.
            const pages = layoutPages(
                [para(lines(6), { borders: boxOf(2) })], PAGE,
                { widowOrphanControl: false });
            const [second] = pages[1]!.paragraphBorders;

            expect(second!.topPx).toBe(pages[1]!.lines[0]!.yPx - 1);
            expect(second!.topPx).toBeLessThan(pages[0]!.paragraphBorders[0]!.topPx);
        });

        it('follows an indent, so the box hugs the text and not the margin', () => {
            const page = layoutPages(
                [para('xx', { borders: boxOf(2), indentLeftPx: 12, indentRightPx: 8 })],
                PAGE)[0]!;
            const [box] = page.paragraphBorders;

            expect(box!.leftPx).toBe(CONTENT_LEFT + 12 - 1);
            expect(box!.rightPx).toBe(CONTENT_RIGHT - 8 + 1);
        });

        it('rules between the paragraphs of a merged box, and takes its room', () => {
            const borders: BoxBorders = { ...boxOf(2), insideH: side(6) };
            const page = layoutPages(
                [para('a', { borders }), para('b', { borders })], PAGE)[0]!;
            const [box] = page.paragraphBorders;

            expect(page.paragraphBorders).toHaveLength(1);
            expect(box!.innerYPx).toHaveLength(1);
            // Six points of rule between two lines that would otherwise abut,
            // drawn down the middle of the gap it opened.
            expect(page.lines[1]!.yPx - page.lines[0]!.yPx).toBe(plainStep() + 6);
            expect(box!.innerYPx[0]).toBe(
                page.lines[0]!.yPx + page.lines[0]!.heightPx + 3);
        });

        it('rules once for each join, not once for each paragraph', () => {
            const borders: BoxBorders = { ...boxOf(2), insideH: side(6) };
            const page = layoutPages(
                [para('a', { borders }), para('b', { borders }), para('c', { borders })],
                PAGE)[0]!;

            expect(page.paragraphBorders[0]!.innerYPx).toHaveLength(2);
        });

        it('counts w:between when deciding whether two boxes are the same', () => {
            // One asks for a rule below it and the other does not: different
            // borders, so two boxes rather than one with a rule through it.
            const page = layoutPages([
                para('a', { borders: { ...boxOf(2), insideH: side(6) } }),
                para('b', { borders: boxOf(2) }),
            ], PAGE)[0]!;

            expect(page.paragraphBorders).toHaveLength(2);
        });

        it('starts the next box with no rules carried over from the last', () => {
            // A ruled run, then a differently bordered one. The second box is
            // its own: the joins belong to the run that made them.
            const ruled: BoxBorders = { ...boxOf(2), insideH: side(6) };
            const page = layoutPages([
                para('a', { borders: ruled }),
                para('b', { borders: ruled }),
                para('c', { borders: boxOf(4) }),
                para('d', { borders: boxOf(4) }),
            ], PAGE)[0]!;

            expect(page.paragraphBorders.map((box) => box.innerYPx.length)).toEqual([1, 0]);
        });

        it('leaves a box with no w:between free of inner rules', () => {
            const borders = boxOf(2);
            const page = layoutPages(
                [para('a', { borders }), para('b', { borders })], PAGE)[0]!;

            expect(page.paragraphBorders[0]!.innerYPx).toEqual([]);
        });

        it('leaves pages with no bordered paragraph carrying an empty list', () => {
            const [page] = layoutPages([para('xx')], PAGE);

            expect(page!.paragraphBorders).toEqual([]);
        });
    });
});

describe('placePageBorder', () => {
    // The measured fixtures are a square 567 twips all round, which cannot
    // tell an edge measured from its own margin from one measured off any
    // other. These margins are all different on purpose.
    const UNEVEN: PageGeometry = {
        widthPx: 1000,
        heightPx: 2000,
        marginTopPx: 10,
        marginRightPx: 20,
        marginBottomPx: 30,
        marginLeftPx: 40,
    };

    const side = (widthPx: number, spacePx: number): BorderSide => ({
        widthPx,
        style: 'single' as BorderStyle,
        colorHex: '#000000',
        ...(0 === spacePx ? {} : { spacePx }),
    });

    const boxOf = (widthPx = 4, spacePx = 6): BoxBorders => ({
        top: side(widthPx, spacePx),
        left: side(widthPx, spacePx),
        bottom: side(widthPx, spacePx),
        right: side(widthPx, spacePx),
    });

    it('measures a `page` border in from each edge of the PAPER', () => {
        // The gap is clear paper, so the rule's centre is half a width further
        // in: six points of space and a four-point rule put it at eight.
        const box = placePageBorder(
            { borders: boxOf(), offsetFrom: 'page' }, UNEVEN);

        expect(box.leftPx).toBe(8);
        expect(box.topPx).toBe(8);
        expect(box.rightPx).toBe(1000 - 8);
        expect(box.bottomPx).toBe(2000 - 8);
    });

    it('measures a `text` border out from each edge’s OWN margin', () => {
        const box = placePageBorder(
            { borders: boxOf(), offsetFrom: 'text' }, UNEVEN);

        expect(box.leftPx).toBe(40 - 8);
        expect(box.topPx).toBe(10 - 8);
        expect(box.rightPx).toBe(1000 - 20 + 8);
        expect(box.bottomPx).toBe(2000 - 30 + 8);
    });

    it('gives a side the page did not ask for no gap of its own', () => {
        const box = placePageBorder(
            { borders: { top: side(4, 6) }, offsetFrom: 'text' }, UNEVEN);

        expect(box.topPx).toBe(10 - 8);
        // No left rule, so nothing stands off the left margin.
        expect(box.leftPx).toBe(40);
    });

    it('is a box that opens and closes on its own page', () => {
        // Unlike a paragraph's, a page border never carries overleaf: the
        // page it is drawn on is the whole of it.
        const box = placePageBorder({ borders: boxOf(), offsetFrom: 'page' }, UNEVEN);

        expect([box.opensHere, box.closesHere]).toEqual([true, true]);
        expect(box.innerYPx).toEqual([]);
    });
});
