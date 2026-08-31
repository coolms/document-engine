import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import type { BorderStyle } from '../../src/layout/borders.js';
import {
    DEFAULT_PARAGRAPH_STYLE,
    fitTable,
    layoutPages,
    verticalSpans,
    type Block,
    type PageGeometry,
    type Paragraph,
    type Table,
    type CellVerticalAlign,
    type TableCell,
    type TableRow,
} from '../../src/layout/page-layout.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MONO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Regular.ttf'))));

const CELL = MONO.measureAdvance('x', 16).widthPt;
const LINE = 20;

/** A page holding exactly 5 lines, 20 characters wide. */
const PAGE: PageGeometry = {
    widthPx: CELL * 20 + 20,
    heightPx: LINE * 5 + 20,
    marginTopPx: 10,
    marginRightPx: 10,
    marginBottomPx: 10,
    marginLeftPx: 10,
};

function para(text: string): Paragraph {
    return {
        runs: [{ text, font: MONO, sizePx: 16 }],
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
    };
}

function cell(text: string, gridSpan = 1): TableCell {
    return { paragraphs: [para(text)], gridSpan };
}

function row(cells: TableCell[], isHeader = false): TableRow {
    return { cells, isHeader, cantSplit: false };
}

/** Two columns of ten monospace cells each, no padding unless asked for. */
function table(rows: TableRow[], overrides: Partial<Table> = {}): Table {
    return {
        rows,
        columnWidthsPx: [CELL * 10, CELL * 10],
        cellMarginLeftPx: 0,
        cellMarginRightPx: 0,
        cellMarginTopPx: 0,
        cellMarginBottomPx: 0,
        spaceBeforePx: 0,
        spaceAfterPx: 0,
        pageBreakBefore: false,
        ...overrides,
    };
}

function lay(blocks: Block[], geometry = PAGE) {
    return layoutPages(blocks, geometry, { widowOrphanControl: false });
}

function textOf(cellLines: readonly { line: { pieces: readonly { text: string }[] } }[]): string {
    return cellLines.flatMap((placed) => placed.line.pieces.map((piece) => piece.text)).join('');
}

describe('table layout', () => {
    describe('rows and cells', () => {
        it('places a row per row and a cell per cell', () => {
            const [page] = lay([table([row([cell('a'), cell('b')]), row([cell('c'), cell('d')])])]);

            expect(page!.rows.length).toBe(2);
            expect(page!.rows[0]!.cells.length).toBe(2);
            expect(page!.rows.map((r) => r.rowIndex)).toEqual([0, 1]);
        });

        it('takes a row\'s height from its TALLEST cell', () => {
            // One cell of one line beside one of three. Sizing the row from the
            // first would put its bottom border through the other's text.
            //
            // Three separate words, not thirty x's: an unbreakable stretch wider
            // than the column is placed alone and allowed to overflow, so it
            // would be ONE line and the row would look correctly sized.
            const tall = { paragraphs: [para('xxxxxxxxxx xxxxxxxxxx xxxxxxxxxx')], gridSpan: 1 };
            const [page] = lay([table([row([cell('short'), tall])])]);

            expect(page!.rows[0]!.heightPx).toBe(LINE * 3);
        });

        it('wraps a cell\'s text at the COLUMN width, not the page width', () => {
            // The control is the point: the same text as an ordinary paragraph
            // has the page's twenty characters and fits on ONE line. In a
            // ten-character column it takes three.
            const text = 'aaaaa bbbbb ccccc';
            const [page] = lay([table([row([cell(text)])])]);
            const [control] = lay([para(text)]);

            expect(control!.lines.length).toBe(1);
            expect(page!.rows[0]!.cells[0]!.lines.length).toBe(3);
            expect(textOf(page!.rows[0]!.cells[0]!.lines.slice(0, 1))).toBe('aaaaa ');
        });

        it('puts each column at its own offset, from the GRID', () => {
            const [page] = lay([table([row([cell('a'), cell('b')])])]);
            const [first, second] = page!.rows[0]!.cells;

            expect(first!.xPx).toBe(PAGE.marginLeftPx);
            expect(second!.xPx).toBe(PAGE.marginLeftPx + CELL * 10);
            expect(first!.widthPx).toBe(CELL * 10);
        });

        it('lets a cell span several grid columns', () => {
            // w:gridSpan. A cell that spans two columns is as wide as both, and
            // the cell after it starts where the SPAN ended.
            const [page] = lay([table([
                row([cell('wide', 2)]),
                row([cell('a'), cell('b')]),
            ], { columnWidthsPx: [CELL * 6, CELL * 6, CELL * 6] })]);

            expect(page!.rows[0]!.cells[0]!.widthPx).toBe(CELL * 12);
            expect(page!.rows[1]!.cells[1]!.xPx).toBe(PAGE.marginLeftPx + CELL * 6);
        });

        it('positions cell text absolutely on the page', () => {
            // A cell lays its text out from its own origin; what comes back has
            // to be where that text actually is.
            const [page] = lay([table([row([cell('a'), cell('b')])])]);
            const second = page!.rows[0]!.cells[1]!;

            expect(second.lines[0]!.xPx).toBe(PAGE.marginLeftPx + CELL * 10);
            expect(second.lines[0]!.yPx).toBe(PAGE.marginTopPx);
        });
    });

    describe('cell padding', () => {
        it('narrows the text column and makes the row taller', () => {
            // 'aaa bbb' is seven characters: it fits the bare ten-character
            // column on one line and cannot fit the six that are left once the
            // padding is taken out. Anything wider would wrap either way and
            // the padding would make no observable difference.
            const text = 'aaa bbb';
            const padded = table([row([cell(text)])], {
                cellMarginLeftPx: CELL * 2,
                cellMarginRightPx: CELL * 2,
                cellMarginTopPx: 5,
                cellMarginBottomPx: 5,
            });

            const [bare] = lay([table([row([cell(text)])])]);
            const [page] = lay([padded]);
            const only = page!.rows[0]!.cells[0]!;

            expect(bare!.rows[0]!.cells[0]!.lines.length).toBe(1);
            expect(only.lines.length).toBe(2);
            expect(page!.rows[0]!.heightPx).toBe(LINE * 2 + 10);
            expect(only.lines[0]!.xPx).toBe(PAGE.marginLeftPx + CELL * 2);
            expect(only.lines[0]!.yPx).toBe(PAGE.marginTopPx + 5);
        });

        it('honours a cell paragraph\'s own spacing and indents', () => {
            // A cell is a little page: what a paragraph does outside a table it
            // does inside one, and ignoring either makes the row the wrong
            // height or the text the wrong width.
            const spaced: Paragraph = {
                runs: [{ text: 'aaa bbb', font: MONO, sizePx: 16 }],
                style: {
                    ...DEFAULT_PARAGRAPH_STYLE,
                    lineHeightPx: LINE,
                    spaceBeforePx: 7,
                    spaceAfterPx: 11,
                    indentLeftPx: CELL * 4,
                },
            };

            const [page] = lay([table([row([{ paragraphs: [spaced], gridSpan: 1 }])])]);
            const only = page!.rows[0]!.cells[0]!;

            // Six characters left after the indent, so the words split.
            expect(only.lines.length).toBe(2);
            expect(only.lines[0]!.xPx).toBe(PAGE.marginLeftPx + CELL * 4);
            expect(only.lines[0]!.yPx).toBe(PAGE.marginTopPx + 7);
            expect(page!.rows[0]!.heightPx).toBe(7 + LINE * 2 + 11);
        });
    });

    describe('breaking across pages', () => {
        it('moves a row that does not fit onto the next page, WHOLE', () => {
            // Five rows of one line each fit; the sixth starts page two.
            const rows = Array.from({ length: 6 }, (_unused, index) => row([cell(`r${index}`)]));
            const pages = lay([table(rows)]);

            expect(pages.length).toBe(2);
            expect(pages[0]!.rows.length).toBe(5);
            expect(pages[1]!.rows.length).toBe(1);
            expect(pages[1]!.rows[0]!.yPx).toBe(PAGE.marginTopPx);
        });

        it('SPLITS a row to fill the page, and carries the rest over', () => {
            // Four lines used, one left, and a two-line row to place. Measured
            // against LibreOffice: the line that fits stays and the other goes
            // on, where moving the whole row would leave the page a line short.
            const pages = lay([
                para('a\nb\nc\nd'),
                table([row([cell('xxxxxxxxxx xxxxxxxxxx')])]),
            ]);

            expect(pages[0]!.rows.length).toBe(1);
            expect(pages[0]!.rows[0]!.heightPx).toBe(LINE);
            expect(textOf(pages[0]!.rows[0]!.cells[0]!.lines)).toBe('xxxxxxxxxx ');
            expect(pages[1]!.rows[0]!.heightPx).toBe(LINE);
            expect(textOf(pages[1]!.rows[0]!.cells[0]!.lines)).toBe('xxxxxxxxxx');
        });

        it('draws no rule across the break it was cut at', () => {
            // The halves meet at a page boundary, so the edge between them is
            // not an edge of the table — a rule there is a line the document
            // never asked for.
            const side = { widthPx: 2, colorHex: '#000000', style: 'solid' as const };
            const pages = lay([
                para('a\nb\nc\nd'),
                table([row([{
                    ...cell('xxxxxxxxxx xxxxxxxxxx'),
                    borders: { top: side, bottom: side },
                }])]),
            ]);

            expect(pages[0]!.rows[0]!.cells[0]!.borders?.top?.widthPx).toBe(2);
            expect(pages[0]!.rows[0]!.cells[0]!.borders?.bottom).toBeUndefined();
            expect(pages[1]!.rows[0]!.cells[0]!.borders?.top).toBeUndefined();
            expect(pages[1]!.rows[0]!.cells[0]!.borders?.bottom?.widthPx).toBe(2);
        });

        it('starts the carried-over lines at the TOP of their new cell', () => {
            // The lines keep the y they had in the whole row unless they are
            // re-based, which would draw the remainder as far down the second
            // page as it had been down the first.
            const pages = lay([
                para('a\nb\nc\nd'),
                table([row([cell('xxxxxxxxxx xxxxxxxxxx')])]),
            ]);
            const carried = pages[1]!.rows[0]!;

            expect(carried.cells[0]!.lines[0]!.yPx).toBe(carried.yPx);
        });

        it('counts the padding above the text when it chooses where to cut', () => {
            // The padding is part of what fills the row. Cutting as though the
            // text began at the row's own top fits one line too many.
            const rows = [row([cell('xxxxxxxxxx xxxxxxxxxx xxxxxxxxxx')])];
            const linesOn = (cellMarginTopPx: number): number => lay([
                para('a\nb'),
                table(rows, { cellMarginTopPx }),
            ])[0]!.rows[0]!.cells[0]!.lines.length;

            // Three lines of room, so all three fit bare; with padding above
            // them only two do.
            expect(linesOn(0)).toBe(3);
            expect(linesOn(LINE)).toBe(2);
        });

        it('moves a row whole when its own height is what does not fit', () => {
            // A `w:trHeight` taller than the text: every line fits the room
            // left, and the ROW still does not. There is nothing to cut, so
            // cutting would leave an empty half behind.
            const pages = lay([
                para('a\nb\nc\nd'),
                table([{ ...row([cell('x')]), heightPx: LINE * 3 }]),
            ]);

            expect(pages[0]!.rows.length).toBe(0);
            expect(pages[1]!.rows[0]!.heightPx).toBe(LINE * 3);
        });

        it('moves a row whole when it says w:cantSplit', () => {
            const pages = lay([
                para('a\nb\nc\nd'),
                table([{
                    ...row([cell('xxxxxxxxxx xxxxxxxxxx')]),
                    cantSplit: true,
                }]),
            ]);

            expect(pages[0]!.rows.length).toBe(0);
            expect(pages[1]!.rows[0]!.heightPx).toBe(LINE * 2);
        });

        it('never cuts a row a merge reaches through', () => {
            // A merged cell is one box over several rows; cutting one of them
            // would leave the box drawn across a page it does not reach.
            const merged = {
                ...cell('m'),
                verticalMerge: 'restart' as const,
            };
            const pages = lay([
                para('a\nb\nc\nd'),
                table([
                    row([merged, cell('xxxxxxxxxx xxxxxxxxxx')]),
                    row([{ ...cell(''), verticalMerge: 'continue' as const }, cell('z')]),
                ]),
            ]);

            expect(pages[0]!.rows.length).toBe(0);
        });

        it('breaks a row taller than a whole page across pages', () => {
            // Something CAN be done for it: it is cut, and cut again, until
            // what is left fits. Measured against LibreOffice
            // (`table-split-row-taller.docx`) — a 150-line row spans three
            // pages there rather than running off the foot of one.
            //
            // This test asserted a single page for as long as the engine ran
            // such a row off the bottom of it.
            const pages = lay([table([row([cell('x '.repeat(40))])])]);

            expect(pages.length).toBeGreaterThan(1);
            expect(pages[0]!.rows.length).toBe(1);
        });

        it('continues after the table where the table ended', () => {
            const pages = lay([
                table([row([cell('a')]), row([cell('b')])]),
                para('after'),
            ]);

            expect(pages[0]!.lines[0]!.yPx).toBe(PAGE.marginTopPx + LINE * 2);
        });

        it('counts a table as content, so a page break after one is not blank', () => {
            // Emptiness that only looked at lines would call a page holding a
            // table empty, and the break would open a blank sheet.
            const pages = lay([
                table([row([cell('a')])]),
                { ...para('next'), style: { ...para('next').style, pageBreakBefore: true } },
            ]);

            expect(pages.length).toBe(2);
            expect(pages[0]!.rows.length).toBe(1);
        });
    });

    describe('repeating header rows', () => {
        const headed = (dataRows: number): Table => table([
            row([cell('HEADER')], true),
            ...Array.from({ length: dataRows }, (_unused, index) => row([cell(`r${index}`)])),
        ]);

        it('repeats the header at the top of the continuation', () => {
            // Header plus four data rows fill page one; the fifth goes over,
            // and arrives under a header of its own.
            const pages = lay([headed(6)]);

            expect(pages.length).toBe(2);
            expect(textOf(pages[1]!.rows[0]!.cells[0]!.lines)).toBe('HEADER');
            expect(pages[1]!.rows[0]!.repeated).toBe(true);
            expect(pages[1]!.rows[0]!.rowIndex).toBe(0);
        });

        it('does not repeat the header in front of itself', () => {
            const pages = lay([headed(6)]);

            expect(pages[0]!.rows.filter((r) => r.repeated).length).toBe(0);
            expect(pages[0]!.rows[0]!.repeated).toBe(false);
        });

        it('carries the FIRST header over with the second', () => {
            // Two header rows, pushed down so the second one overflows.
            //
            // Measured (`table-two-headers.docx`): the whole table
            // moves. LibreOffice ended the page with its filler and opened the
            // next with H1, H2 and the body row — it did not leave H1 standing
            // alone above a page break. This test used to assert that it did.
            const pages = lay([
                para('a\nb\nc\nd'),
                table([
                    row([cell('H1')], true),
                    row([cell('H2')], true),
                    row([cell('data')]),
                ]),
            ]);

            expect(textOf(pages[1]!.rows[0]!.cells[0]!.lines)).toBe('H1');
            expect(textOf(pages[1]!.rows[1]!.cells[0]!.lines)).toBe('H2');
            expect(pages[1]!.rows[0]!.repeated).toBe(false);
            expect(pages[0]!.rows).toEqual([]);
        });

        it('marks a repeat so a renderer can tell it from the original', () => {
            const pages = lay([headed(6)]);
            const repeats = pages.flatMap((page) => page.rows).filter((r) => r.repeated);

            expect(repeats.length).toBe(1);
        });

        it('treats only LEADING rows as headers', () => {
            // A w:tblHeader in the middle of a table is not a header, and
            // repeating it would move content that belongs further down.
            const pages = lay([table([
                row([cell('a')]),
                row([cell('MIDDLE')], true),
                ...Array.from({ length: 6 }, () => row([cell('x')])),
            ])]);

            expect(pages.flatMap((page) => page.rows).filter((r) => r.repeated)).toEqual([]);
        });

        it('repeats the header above the REST of a row it cut', () => {
            // Measured against LibreOffice: a three-line header and three-line
            // rows on a fourteen-line page put two lines of the fourth row at
            // the foot, then repeated the header and carried the third line
            // over beneath it.
            const threeLines = 'xxxxxxxxxx xxxxxxxxxx xxxxxxxxxx';
            const pages = lay([table([
                row([cell(threeLines)], true),
                ...Array.from({ length: 4 }, () => row([cell(threeLines)])),
            ])]);

            expect(pages[1]!.rows[0]!.repeated).toBe(true);
            expect(pages[1]!.rows[1]!.repeated).toBe(false);
        });

        it('skips the repeat when the header alone fills the page', () => {
            // Repeating a header with no room for any of the row it heads
            // would push the table forward a row at a time and never finish.
            const fiveLines = 'xxxxxxxxxx '.repeat(5);
            const pages = lay([table([
                row([cell(fiveLines)], true),
                ...Array.from({ length: 3 }, () => row([cell('a')])),
            ])]);

            expect(pages.flatMap((page) => page.rows).filter((r) => r.repeated)).toEqual([]);
        });
    });

    describe('provenance', () => {
        it('records which block and which row each placed row came from', () => {
            const pages = lay([para('before'), table([row([cell('a')]), row([cell('b')])])]);

            expect(pages[0]!.rows.map((r) => [r.blockIndex, r.rowIndex])).toEqual([[1, 0], [1, 1]]);
        });
    });
});

describe('vertical alignment and row height', () => {
    /** A cell of several one-line paragraphs, so its height is countable. */
    const stack = (...texts: string[]): TableCell => ({
        paragraphs: texts.map(para),
        gridSpan: 1,
    });

    const aligned = (align: CellVerticalAlign | undefined): TableCell => ({
        ...cell('x'),
        ...(undefined === align ? {} : { verticalAlign: align }),
    });

    /** Where each cell of the first row drew its first line, top of page aside. */
    const tops = (blocks: Block[]): number[] => {
        const [page] = lay(blocks, PAGE);
        const placed = page!.rows[0]!;

        return placed.cells.map((held) => held.lines[0]!.yPx - placed.yPx);
    };

    it('leaves a cell at the top of a row taller than it', () => {
        // Three lines beside one. Without an alignment the short cell starts
        // where the tall one does.
        expect(tops([table([row([stack('a', 'b', 'c'), aligned(undefined)])])]))
            .toEqual([0, 0]);
        expect(tops([table([row([stack('a', 'b', 'c'), aligned('top')])])]))
            .toEqual([0, 0]);
    });

    it('centres and bottoms a cell in the room the row has spare', () => {
        // Two lines of slack: centred takes one, bottomed takes both.
        expect(tops([table([row([stack('a', 'b', 'c'), aligned('center')])])]))
            .toEqual([0, LINE]);
        expect(tops([table([row([stack('a', 'b', 'c'), aligned('bottom')])])]))
            .toEqual([0, LINE * 2]);
    });

    it('moves the whole cell, not only its first line', () => {
        // Vertical alignment is one shift of a block of paragraphs, which is
        // what tells it apart from justifying them to fill the cell.
        const [page] = lay(
            [table([row([stack('a', 'b', 'c', 'd'), { ...stack('x', 'y'), verticalAlign: 'bottom' }])])],
            PAGE,
        );
        const placed = page!.rows[0]!;

        expect(placed.cells[1]!.lines.map((line) => line.yPx - placed.yPx))
            .toEqual([LINE * 2, LINE * 3]);
    });

    it('holds a row open to the height it asks for', () => {
        const [page] = lay(
            [table([{ ...row([cell('a'), cell('b')]), heightPx: LINE * 3 }])],
            PAGE,
        );

        expect(page!.rows[0]!.heightPx).toBe(LINE * 3);
    });

    it('reads a stated height as a MINIMUM unless the rule says exact', () => {
        // Measured against LibreOffice: a row asking for less than its text
        // needs grows to the text under every rule but `exact`.
        const tall = [stack('a', 'b', 'c'), cell('b')];

        expect(lay(
            [table([{ ...row(tall), heightPx: LINE }])], PAGE,
        )[0]!.rows[0]!.heightPx).toBe(LINE * 3);
        expect(lay(
            [table([{ ...row(tall), heightPx: LINE, heightRule: 'atLeast' }])], PAGE,
        )[0]!.rows[0]!.heightPx).toBe(LINE * 3);
    });

    it('lets an exact row be SHORTER than its own text, and drops what hangs out', () => {
        // LibreOffice drew the second of three lines, whose top was still
        // inside the row, and dropped the third.
        const [page] = lay(
            [table([
                { ...row([stack('a', 'b', 'c'), cell('z')]), heightPx: LINE + 1, heightRule: 'exact' },
                row([cell('next'), cell('n')]),
            ])],
            PAGE,
        );
        const placed = page!.rows[0]!;

        expect(placed.heightPx).toBe(LINE + 1);
        expect(placed.cells[0]!.lines.map((line) => line.line.pieces[0]!.text))
            .toEqual(['a', 'b']);
        // ...and the row below it starts at the height the row was given, not
        // at the foot of the text hanging out of it.
        expect(page!.rows[1]!.yPx - placed.yPx).toBe(LINE + 1);
    });

    it('measures the overhang from below the cell\'s TOP PADDING', () => {
        // The padding is part of what fills the row: a line sitting one pixel
        // inside a bare row hangs out of a padded one. Every fixture with no
        // padding at all agrees whichever way this is counted.
        const rows = [{
            ...row([stack('a', 'b'), cell('z')]),
            heightPx: LINE + 1,
            heightRule: 'exact' as const,
        }];
        const drawn = (marginTopPx: number): string[] => lay(
            [table(rows, { cellMarginTopPx: marginTopPx })], PAGE,
        )[0]!.rows[0]!.cells[0]!.lines.map((line) => line.line.pieces[0]!.text);

        expect(drawn(0)).toEqual(['a', 'b']);
        expect(drawn(2)).toEqual(['a']);
    });

    it('does not align a cell in a row too short for it', () => {
        // There is no spare room to share out, and pulling the text UP by the
        // shortfall would hide the first line instead of the last.
        const [page] = lay(
            [table([{
                ...row([{ ...stack('a', 'b', 'c'), verticalAlign: 'bottom' }, cell('z')]),
                heightPx: LINE,
                heightRule: 'exact',
            }])],
            PAGE,
        );

        expect(page!.rows[0]!.cells[0]!.lines[0]!.yPx - page!.rows[0]!.yPx).toBe(0);
    });
});

describe('vertically merged cells', () => {
    const stack = (...texts: string[]): TableCell => ({
        paragraphs: texts.map(para), gridSpan: 1,
    });
    const merged = (text: string, extra: Partial<TableCell> = {}): TableCell => ({
        ...cell(text), verticalMerge: 'restart', ...extra,
    });
    const swallowed = (text = ''): TableCell => ({
        ...cell(text), verticalMerge: 'continue',
    });

    describe('working out the spans', () => {
        it('counts a restart down to its last continuation', () => {
            expect(verticalSpans(table([
                row([merged('m'), cell('a')]),
                row([swallowed(), cell('b')]),
                row([swallowed(), cell('c')]),
            ]))).toEqual([[3, 1], [0, 1], [0, 1]]);
        });

        it('closes a merge at the first row that does not continue it', () => {
            expect(verticalSpans(table([
                row([merged('m'), cell('a')]),
                row([swallowed(), cell('b')]),
                row([cell('own'), cell('c')]),
                row([swallowed(), cell('d')]),
            ]))).toEqual([[2, 1], [0, 1], [1, 1], [1, 1]]);
        });

        it('matches a merge by its grid COLUMN, not by its place in the row', () => {
            // The first row's opening cell covers two columns, so the cell
            // beside it starts at column two — which is where the merge below
            // it must be looked for. Counting cells instead would pair the
            // merge with the wrong column entirely.
            expect(verticalSpans(table([
                row([cell('wide', 2), merged('m')]),
                row([cell('a'), cell('b'), swallowed()]),
            ], { columnWidthsPx: [CELL * 5, CELL * 5, CELL * 5] })))
                .toEqual([[1, 2], [1, 1, 0]]);
        });

        it('treats a continuation with nothing above it as an ordinary cell', () => {
            expect(verticalSpans(table([row([swallowed('a'), cell('b')])])))
                .toEqual([[1, 1]]);
        });
    });

    describe('laying them out', () => {
        const rowsOf = (...rows: TableRow[]) => lay([table(rows)])[0]!.rows;

        it('draws the merged cell once, over the whole span', () => {
            const placed = rowsOf(
                row([merged('m'), cell('a')]),
                row([swallowed(), cell('b')]),
                row([swallowed(), cell('c')]),
            );

            expect(placed[0]!.cells[0]!.heightPx).toBe(LINE * 3);
            // ...and the rows below hold only their own cell, so nothing draws
            // a border through the middle of it.
            expect(placed[1]!.cells.length).toBe(1);
            expect(placed[2]!.cells.length).toBe(1);
        });

        it('leaves the rows their own heights', () => {
            const placed = rowsOf(
                row([merged('m'), cell('a')]),
                row([swallowed(), stack('b', 'c')]),
            );

            expect(placed.map((placedRow) => placedRow.heightPx)).toEqual([LINE, LINE * 2]);
        });

        it('grows the LAST row of the span for a merge taller than it', () => {
            // Measured against LibreOffice: five lines merged across two
            // one-line rows left the first at one line and made the second
            // four, so the span comes to five.
            const placed = rowsOf(
                row([
                    { ...stack('t1', 't2', 't3', 't4', 't5'), verticalMerge: 'restart' },
                    cell('a'),
                ]),
                row([swallowed(), cell('b')]),
            );
            const [first] = placed;

            expect(placed.map((placedRow) => placedRow.heightPx)).toEqual([LINE, LINE * 4]);
            expect(first!.cells[0]!.heightPx).toBe(LINE * 5);
        });

        it('ignores what a swallowed cell says', () => {
            // Word's rule, and LibreOffice drew none of it.
            const placed = rowsOf(
                row([merged('KEEP'), cell('a')]),
                row([swallowed('DROPPED'), cell('b')]),
            );

            expect(placed.flatMap((placedRow) => placedRow.cells)
                .flatMap((held) => held.lines)
                .map((line) => line.line.pieces[0]!.text))
                .toEqual(['KEEP', 'a', 'b']);
        });

        it('aligns a merged cell over the SPAN, not over its own row', () => {
            const placed = rowsOf(
                row([merged('m', { verticalAlign: 'center' }), cell('a')]),
                row([swallowed(), cell('b')]),
                row([swallowed(), cell('c')]),
            );

            expect(placed[0]!.cells[0]!.lines[0]!.yPx - placed[0]!.yPx).toBe(LINE);
        });

        it('keeps the rows a merge covers on one page', () => {
            // Four ordinary rows fill all but one line of the page, so the
            // pair would be split at exactly the boundary the merge forbids.
            const placed = lay([table([
                row([cell('a'), cell('a')]),
                row([cell('b'), cell('b')]),
                row([cell('c'), cell('c')]),
                row([cell('d'), cell('d')]),
                row([merged('m'), cell('e')]),
                row([swallowed(), cell('f')]),
            ])]);

            expect(placed.map((page) => page.rows.map((held) => held.rowIndex)))
                .toEqual([[0, 1, 2, 3], [4, 5]]);
        });
    });

    describe('where a merge meets the rest of the table', () => {
        const side = (widthPx: number) => ({ widthPx, colorHex: '#000000', style: 'solid' as const });

        it('closes a merge that a row skips over entirely', () => {
            // The middle row has one cell, so it says nothing about the second
            // column. A merge cannot reach across a row that does not carry it,
            // and the continuation below is then a cell of its own.
            expect(verticalSpans(table([
                row([cell('a'), merged('m')]),
                row([cell('only')]),
                row([cell('b'), swallowed()]),
            ]))).toEqual([[1, 1], [1], [1, 1]]);
        });

        it('settles the merge that ends SOONEST first', () => {
            // Six lines merged down all four rows, four lines merged down the
            // middle two. Growing the outer one first would push its shortfall
            // onto the last row and then grow the inner one on top of that,
            // making the table eight lines tall. LibreOffice printed six.
            const rows = [
                row([
                    { ...stack('x1', 'x2', 'x3', 'x4', 'x5', 'x6'), verticalMerge: 'restart' as const },
                    cell('a'),
                ]),
                row([swallowed(), { ...stack('y1', 'y2', 'y3', 'y4'), verticalMerge: 'restart' as const }]),
                row([swallowed(), swallowed()]),
                row([swallowed(), cell('d')]),
            ];
            const placed = lay([table(rows)], { ...PAGE, heightPx: LINE * 20 })[0]!.rows;

            // The middle rows carry nothing but merge, so their own height is
            // nil and the inner merge's shortfall all lands on the second of
            // them. What matters is the TOTAL: settling the outer merge first
            // would grow the last row too and make the table ten lines.
            expect(placed.map((held) => held.heightPx)).toEqual([LINE, 0, LINE * 4, LINE]);
            expect(placed.reduce((sum, held) => sum + held.heightPx, 0)).toBe(LINE * 6);
            expect(placed[0]!.cells[0]!.heightPx).toBe(LINE * 6);
        });

        it('carries a merged cell\'s edge to the row below the SPAN', () => {
            // The row below the span declares a thinner edge, and the two have
            // to agree or the join draws twice. The row below the merge's FIRST
            // row is not that row — nothing of the merge ends there.
            const placed = lay([table([
                row([{ ...merged('m'), borders: { bottom: side(4) } }, cell('a')]),
                row([swallowed(), cell('b')]),
                row([{ ...cell('c'), borders: { top: side(1) } }, cell('d')]),
            ])])[0]!.rows;

            expect(placed[0]!.cells[0]!.borders?.bottom?.widthPx).toBe(4);
            expect(placed[2]!.cells[0]!.borders?.top?.widthPx).toBe(4);
        });

        it('places a merge taller than the page whole rather than splitting it', () => {
            // Eight merged lines will not fit a five-line page under any
            // arrangement. Breaking inside the span would draw the top of the
            // cell on one page and its foot below the bottom of another.
            const placed = lay([table([
                row([
                    { ...stack('m1', 'm2', 'm3', 'm4', 'm5', 'm6', 'm7', 'm8'), verticalMerge: 'restart' as const },
                    cell('a'),
                ]),
                row([swallowed(), cell('b')]),
            ])]);

            expect(placed.map((page) => page.rows.map((held) => held.rowIndex))).toEqual([[0, 1]]);
        });
    });
});

describe('a table nested inside a cell', () => {
    const empty: Paragraph = {
        runs: [{ text: '', font: MONO, sizePx: 16 }],
        style: { ...DEFAULT_PARAGRAPH_STYLE, lineHeightPx: LINE },
    };

    /** A one-row, one-column table half the width of a column. */
    const inner = (text: string): Table => table(
        [row([cell(text)])],
        { columnWidthsPx: [CELL * 5] },
    );

    const holding = (...blocks: Block[]): TableCell => ({ paragraphs: blocks, gridSpan: 1 });

    it('lays the nested table out instead of losing it', () => {
        const [page] = lay([table([row([holding(inner('deep')), cell('side')])])]);
        const outer = page!.rows[0]!;

        expect(outer.cells[0]!.rows.length).toBe(1);
        expect(outer.cells[0]!.rows[0]!.cells[0]!.lines
            .map((line) => line.line.pieces[0]!.text)).toEqual(['deep']);
    });

    it('stacks the blocks around it in document order', () => {
        // Measured against LibreOffice: a line, then two nested rows, then a
        // line — and the outer row four lines tall to hold them.
        const [page] = lay(
            [table([row([holding(para('above'), inner('deep'), para('below')), cell('side')])])],
            { ...PAGE, heightPx: LINE * 20 },
        );
        const held = page!.rows[0]!.cells[0]!;

        expect(page!.rows[0]!.heightPx).toBe(LINE * 3);
        expect(held.lines.map((line) => line.yPx - page!.rows[0]!.yPx)).toEqual([0, LINE * 2]);
        expect(held.rows[0]!.yPx - page!.rows[0]!.yPx).toBe(LINE);
    });

    it('puts the nested table at the holding cell\'s own origin', () => {
        const [page] = lay([table([row([cell('a'), holding(inner('deep'))])])], PAGE);
        const outer = page!.rows[0]!;
        const nested = outer.cells[1]!.rows[0]!;

        expect(nested.cells[0]!.xPx).toBe(outer.cells[1]!.xPx);
        expect(nested.yPx).toBe(outer.yPx);
    });

    it('drops the empty paragraph OOXML makes a cell END with', () => {
        // Structure, not content: LibreOffice drew it a twentieth of a point
        // tall. Giving it a line would make every cell holding a table taller
        // than the document says.
        const [page] = lay([table([row([holding(inner('deep'), empty), cell('side')])])]);

        expect(page!.rows[0]!.heightPx).toBe(LINE);
    });

    it('keeps an empty paragraph that is NOT that one', () => {
        // The same paragraph one place earlier, and one that ends a cell with
        // no table above it, both take a whole line — which is what makes the
        // rule about the table below it rather than about emptiness.
        const after = lay([table([row([holding(inner('deep'), empty, para('x')), cell('s')])])]);
        const alone = lay([table([row([holding(para('x'), empty), cell('s')])])]);

        expect(after[0]!.rows[0]!.heightPx).toBe(LINE * 3);
        expect(alone[0]!.rows[0]!.heightPx).toBe(LINE * 2);
    });

    it('keeps the space a nested table asks for above and below it', () => {
        // A table between two paragraphs carries its own spacing, and a cell
        // that ignored it would close up around the table and come out short.
        const spaced = table([row([cell('deep')])], {
            columnWidthsPx: [CELL * 5],
            spaceBeforePx: 6,
            spaceAfterPx: 9,
        });
        const [page] = lay(
            [table([row([holding(para('above'), spaced, para('below')), cell('side')])])],
            { ...PAGE, heightPx: LINE * 20 },
        );
        const held = page!.rows[0]!;

        expect(held.heightPx).toBe(LINE * 3 + 15);
        expect(held.cells[0]!.rows[0]!.yPx - held.yPx).toBe(LINE + 6);
        expect(held.cells[0]!.lines[1]!.yPx - held.yPx).toBe(LINE * 2 + 15);
    });

    it('nests as many levels down as the document does', () => {
        const twice = table([row([holding(inner('deep'))])], { columnWidthsPx: [CELL * 7] });
        const [page] = lay([table([row([holding(twice), cell('side')])])]);

        const first = page!.rows[0]!.cells[0]!.rows[0]!;
        const second = first.cells[0]!.rows[0]!;

        expect(second.cells[0]!.lines[0]!.line.pieces[0]!.text).toBe('deep');
        // Each level starts where the cell holding it starts.
        expect(second.cells[0]!.xPx).toBe(first.cells[0]!.xPx);
    });

    it('measures the holding cell from the nested table\'s height', () => {
        const tall = table(
            [row([cell('r1')]), row([cell('r2')]), row([cell('r3')])],
            { columnWidthsPx: [CELL * 5] },
        );
        const [page] = lay(
            [table([row([holding(tall), cell('side')])])],
            { ...PAGE, heightPx: LINE * 20 },
        );

        expect(page!.rows[0]!.heightPx).toBe(LINE * 3);
    });
});

describe('fitTable', () => {
    const COLUMN = 1000;

    const table = (over: Partial<Table> = {}): Table => ({
        rows: [],
        columnWidthsPx: [200, 300],
        cellMarginLeftPx: 10,
        cellMarginRightPx: 10,
        cellMarginTopPx: 0,
        cellMarginBottomPx: 0,
        spaceBeforePx: 0,
        spaceAfterPx: 0,
        pageBreakBefore: false,
        ...over,
    });

    it('leaves a table with no declared width on its grid', () => {
        const { table: fitted, offsetPx } = fitTable(table(), COLUMN);

        expect(fitted.columnWidthsPx).toEqual([200, 300]);
        expect(offsetPx).toBe(0);
    });

    it('scales the WHOLE grid to a declared width, not just the last column', () => {
        const { table: fitted } = fitTable(table({ preferredWidthPx: 250 }), COLUMN);

        expect(fitted.columnWidthsPx).toEqual([100, 150]);
    });

    it('resolves a fraction against the column it lands in', () => {
        expect(fitTable(table({ preferredWidthFraction: 0.5 }), COLUMN)
            .table.columnWidthsPx).toEqual([200, 300]);
        expect(fitTable(table({ preferredWidthFraction: 0.5 }), 250)
            .table.columnWidthsPx).toEqual([50, 75]);
    });

    it('prefers the fraction where a table somehow states both', () => {
        const { table: fitted } = fitTable(
            table({ preferredWidthFraction: 0.5, preferredWidthPx: 100 }), COLUMN);

        expect(fitted.columnWidthsPx).toEqual([200, 300]);
    });

    it('centres and right-aligns against the resolved width', () => {
        const half = { preferredWidthFraction: 0.5 } as const;

        expect(fitTable(table({ ...half, alignment: 'center' }), COLUMN).offsetPx).toBe(250);
        expect(fitTable(table({ ...half, alignment: 'right' }), COLUMN).offsetPx).toBe(500);
        expect(fitTable(table({ ...half, alignment: 'left' }), COLUMN).offsetPx).toBe(0);
    });

    it('hangs a left border half outside the column it starts at', () => {
        const bordered = table({
            borders: { left: { widthPx: 6, style: 'single' as BorderStyle, colorHex: '#000000' } },
        });

        expect(fitTable(bordered, COLUMN).offsetPx).toBe(-3);
    });

    it('drops the hang once the table is indented, centred or pushed right', () => {
        const left = { widthPx: 6, style: 'single' as BorderStyle, colorHex: '#000000' };
        const bordered = (over: Partial<Table>): Table =>
            table({ borders: { left }, ...over });

        // The indent counts to the cell's TEXT, so the border falls a cell
        // margin left of the number — and the hang plays no further part.
        expect(fitTable(bordered({ indentPx: 40 }), COLUMN).offsetPx).toBe(30);
        expect(fitTable(bordered({ alignment: 'center' }), COLUMN).offsetPx).toBe(250);
        expect(fitTable(bordered({ alignment: 'right' }), COLUMN).offsetPx).toBe(500);
    });

    it('ignores an indent on a table that is not left-aligned', () => {
        const indented = { indentPx: 40, preferredWidthFraction: 0.5 } as const;

        expect(fitTable(table({ ...indented, alignment: 'center' }), COLUMN).offsetPx)
            .toBe(fitTable(table({ preferredWidthFraction: 0.5, alignment: 'center' }), COLUMN)
                .offsetPx);
    });

    it('falls back to the grid for a width of nought or less', () => {
        expect(fitTable(table({ preferredWidthPx: 0 }), COLUMN).table.columnWidthsPx)
            .toEqual([200, 300]);
        expect(fitTable(table({ preferredWidthFraction: -1 }), COLUMN).table.columnWidthsPx)
            .toEqual([200, 300]);
    });

    it('survives a table with no columns at all', () => {
        const empty = fitTable(table({ columnWidthsPx: [], preferredWidthPx: 200 }), COLUMN);

        expect(empty.table.columnWidthsPx).toEqual([]);
        expect(empty.offsetPx).toBe(0);
    });
});
