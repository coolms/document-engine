import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { FontCatalogue, type FontManifest } from '../../src/word/font-catalogue.js';
import {
    paginateFlow,
    type BlockStyle,
    type FlowBlock,
    type FlowItem,
    type FlowTable,
    type FlowTableCell,
    type FlowTableRow,
} from '../../src/richtext/flow.js';
import { isTable, type Paragraph, type ParagraphStyle, type Table, type TableCell }
    from '../../src/layout/page-layout.js';
import type { BorderSide, BorderStyle, BoxBorders } from '../../src/layout/borders.js';
import type { StyledRun } from '../../src/layout/line-breaker.js';
import type { PageGeometry } from '../../src/layout/page-layout.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MANIFEST = JSON.parse(readFileSync(join(FONT_DIR, 'fonts.manifest.json'), 'utf8')) as FontManifest;
const FONTS = FontCatalogue.load(MANIFEST, (file) => new Uint8Array(readFileSync(join(FONT_DIR, file))));

/** Monospaced, so the arithmetic below is exact rather than approximate. */
const MONO = FONTS.resolve('Courier New', false, false).font;
const CELL = MONO.measureAdvance('x', 16).widthPt;
const LINE = 20;

/** A page holding exactly 5 lines of 20 characters. */
const PAGE: PageGeometry = {
    widthPx: CELL * 20 + 20,
    heightPx: LINE * 5 + 20,
    marginTopPx: 10,
    marginRightPx: 10,
    marginBottomPx: 10,
    marginLeftPx: 10,
};

const BASE = { fontFamily: 'Courier New', fontSizePx: 16, lineHeightPx: LINE };

function flow(blocks: readonly FlowItem[], widowOrphanControl = false) {
    return paginateFlow(blocks, PAGE, { fonts: FONTS, base: BASE, widowOrphanControl });
}

/** The laid-out block at `index`, as a paragraph. */
function paragraphAt(pagination: ReturnType<typeof flow>, index: number): Paragraph {
    const block = pagination.paragraphs[index];
    if (undefined === block || isTable(block)) {
        throw new Error(`block ${index} is not a paragraph`);
    }

    return block;
}

/** A one-paragraph cell at a caller position. */
function tableCell(text: string, at: number, isHeader = false) {
    return {
        blocks: [{ spans: [{ text, at }], at }],
        ...(isHeader ? { isHeader: true } : {}),
    };
}

/** A block of `count` words, each filling one 20-character line. */
function block(count: number, at = 1): FlowBlock {
    return { spans: [{ text: Array.from({ length: count }, () => 'x'.repeat(20)).join(' '), at }], at };
}

describe('paginateFlow', () => {
    describe('page starts', () => {
        it('reports one start per page after the first', () => {
            const { pages, pageStarts } = flow([block(12)]);

            expect(pages.length).toBe(3);
            expect(pageStarts.length).toBe(pages.length - 1);
            expect(pageStarts.map((start) => start.page)).toEqual([1, 2]);
        });

        it('says nothing at all for a document that fits on one page', () => {
            expect(flow([block(3)]).pageStarts).toEqual([]);
        });

        it('points at the first character of the page, in the caller\'s coordinates', () => {
            // Each line is 20 characters plus the space that separated it, so
            // page two starts at offset 5 * 21 = 105 — and at position 1 + 105
            // in the caller's document, because its text began at 1.
            const { pageStarts } = flow([block(12)]);

            expect(pageStarts[0]!.textOffset).toBe(105);
            expect(pageStarts[0]!.at).toBe(106);
            expect(pageStarts[0]!.blockIndex).toBe(0);
        });

        it('distinguishes a break BETWEEN blocks from one inside a block', () => {
            // Between: page two starts at the top of block 1, offset 0.
            const between = flow([block(5, 1), block(3, 200)]);
            expect(between.pageStarts[0]).toEqual({
                page: 1,
                blockIndex: 1,
                textOffset: 0,
                at: 200,
                insideBlock: false,
            });

            // Inside: page two starts part-way through block 0.
            const inside = flow([block(8, 1)]);
            expect(inside.pageStarts[0]!.insideBlock).toBe(true);
            expect(inside.pageStarts[0]!.blockIndex).toBe(0);
        });
    });

    describe('mapping offsets back to the caller\'s positions', () => {
        it('walks the SPANS, so inline content between them is not skipped', () => {
            // The gap between span positions is the editor's own — an image or a
            // field chip occupying a position and no text. Adding the text
            // offset to the block's start would drift past every one of them.
            const blocks: FlowBlock[] = [{
                at: 1,
                spans: [
                    { text: 'x'.repeat(20) + ' ', at: 1 },
                    // 100 positions later: something this engine never sees sits
                    // between the two spans.
                    { text: 'x'.repeat(20) + ' ', at: 121 },
                    { text: 'x'.repeat(20) + ' ', at: 241 },
                    { text: 'x'.repeat(20) + ' ', at: 361 },
                    { text: 'x'.repeat(20) + ' ', at: 481 },
                    { text: 'x'.repeat(20), at: 601 },
                ],
            }];

            const { pageStarts } = flow(blocks);

            // Page two starts at text offset 105, which is the first character
            // of the sixth span — position 601, NOT 1 + 105.
            expect(pageStarts[0]!.textOffset).toBe(105);
            expect(pageStarts[0]!.at).toBe(601);
        });

        it('resolves an offset in the middle of a span', () => {
            const blocks: FlowBlock[] = [{
                at: 1,
                spans: [{ text: Array.from({ length: 12 }, () => 'x'.repeat(20)).join(' '), at: 50 }],
            }];

            expect(flow(blocks).pageStarts[0]!.at).toBe(50 + 105);
        });

        it('reports null rather than guessing when the caller gave no positions', () => {
            // Words, not one long run of x: an unbreakable stretch wider than
            // the column is placed alone and allowed to overflow, so it never
            // wraps and there would be no second page to point at.
            const text = Array.from({ length: 12 }, () => 'x'.repeat(20)).join(' ');
            const blocks: FlowBlock[] = [{ spans: [{ text }] }];

            const { pageStarts } = flow(blocks);

            expect(pageStarts.length).toBeGreaterThan(0);
            expect(pageStarts[0]!.at).toBeNull();
            // The text offset is still known — only the caller's coordinate is not.
            expect(pageStarts[0]!.textOffset).toBe(105);
        });
    });

    describe('reflow, which is the point', () => {
        it('moves overflowing content onto the next page WITHOUT an explicit break', () => {
            // The complaint this whole engine exists for: text that runs past
            // the bottom of a page continues on the next one instead of the page
            // growing to swallow it.
            const { pages } = flow([block(11)]);

            expect(pages.length).toBe(3);
            expect(pages.map((page) => page.lines.length)).toEqual([5, 5, 1]);
        });

        it('honours an explicit break as well', () => {
            const { pages, pageStarts } = flow([
                { spans: [{ text: 'first', at: 1 }], at: 1 },
                { spans: [{ text: 'second', at: 20 }], at: 20, pageBreakBefore: true },
            ]);

            expect(pages.length).toBe(2);
            expect(pageStarts[0]!.at).toBe(20);
            expect(pageStarts[0]!.insideBlock).toBe(false);
        });

        it('applies widow and orphan control when asked', () => {
            // Six lines on a five-line page splits 5/1 without it and 4/2 with,
            // and the editor must show whichever the .docx will do.
            expect(flow([block(6)], false).pages.map((p) => p.lines.length)).toEqual([5, 1]);
            expect(flow([block(6)], true).pages.map((p) => p.lines.length)).toEqual([4, 2]);
        });
    });

    describe('styling', () => {
        it('gives an empty block a real height', () => {
            const { pages } = flow([{ spans: [], at: 1 }]);

            expect(pages[0]!.lines.length).toBe(1);
            expect(pages[0]!.lines[0]!.heightPx).toBe(LINE);
        });

        it('gives an empty block a height from the FONT when none is fixed', () => {
            // The test above cannot see this: an explicit line height applies
            // whether or not the block has a run. Only the font-derived height
            // needs a run to be derived from, and an empty paragraph with no
            // height lets a page hold more than it should.
            const { pages } = paginateFlow([{ spans: [], at: 1 }], PAGE, {
                fonts: FONTS,
                base: { fontFamily: 'Courier New', fontSizePx: 16 },
            });

            expect(pages[0]!.lines[0]!.heightPx).toBeCloseTo(MONO.naturalLineHeight(16), 9);
        });

        it('lets a span override the block\'s font and size', () => {
            const blocks: FlowBlock[] = [{
                at: 1,
                spans: [
                    { text: 'plain ', at: 1 },
                    { text: 'bold', at: 7, bold: true },
                ],
            }];

            const paragraph = paragraphAt(flow(blocks), 0);
            expect(paragraph.runs[0]!.font).not.toBe(paragraph.runs[1]!.font);
        });

        it('takes paragraph spacing into account when filling a page', () => {
            // Spacing is page-filling content: ignoring it fits more lines on a
            // page than the document will.
            const spaced = Array.from({ length: 4 }, (_unused, index): FlowBlock => ({
                spans: [{ text: 'one line', at: 1 + index * 10 }],
                style: { spaceAfterPx: 30 },
            }));

            // Four lines of 20 plus three gaps of 30 is 170, past the 100 the
            // page has.
            expect(flow(spaced).pages.length).toBeGreaterThan(1);
        });
    });

    describe('tables', () => {
        /** Two columns of ten cells each; a header row plus `count` data rows. */
        function grid(count: number, header = false): FlowTable {
            const rows = [];
            let at = 1;

            if (header) {
                rows.push({ cells: [tableCell('HEAD', at, true), tableCell('ING', at + 10, true)] });
                at += 20;
            }
            for (let index = 0; index < count; index++) {
                rows.push({ cells: [tableCell(`r${index}`, at), tableCell('x', at + 10)] });
                at += 20;
            }

            return { rows, columnWidthsPx: [CELL * 10, CELL * 10], at: 1 };
        }

        it('lays a table out as rows on the page', () => {
            const { pages } = flow([grid(3)]);

            expect(pages.length).toBe(1);
            expect(pages[0]!.rows.length).toBe(3);
            expect(pages[0]!.lines.length).toBe(0);
        });

        it('flows rows onto the next page and points at the ROW that starts it', () => {
            // Six one-line rows on a five-line page: the sixth begins page two,
            // and the editor is told which row and where it is.
            const { pages, pageStarts } = flow([grid(6)]);

            expect(pages.length).toBe(2);
            expect(pageStarts.length).toBe(1);
            expect(pageStarts[0]!.rowIndex).toBe(5);
            expect(pageStarts[0]!.blockIndex).toBe(0);
            expect(pageStarts[0]!.at).toBe(1 + 5 * 20);
            expect(pageStarts[0]!.insideBlock).toBe(true);
        });

        it('repeats a header row, and does not point the page start at the repeat', () => {
            // The repeat is drawn again at the top; the page BEGINS at the row
            // after it, and pointing an editor at the repeat would send the
            // caret backwards into content it has already passed.
            const { pages, pageStarts } = flow([grid(6, true)]);

            expect(pages.length).toBe(2);
            expect(pages[1]!.rows[0]!.repeated).toBe(true);
            expect(pageStarts[0]!.rowIndex).toBeGreaterThan(0);
            expect(pages[1]!.rows.find((row) => row.rowIndex === pageStarts[0]!.rowIndex)!.repeated)
                .toBe(false);
        });

        it('treats a row as a header only when EVERY cell is one', () => {
            // One header cell in an ordinary row must not make that row repeat
            // on every page.
            const mixed: FlowTable = {
                at: 1,
                columnWidthsPx: [CELL * 10, CELL * 10],
                rows: [
                    { cells: [tableCell('a', 1, true), tableCell('b', 11)] },
                    ...Array.from({ length: 6 }, (_unused, index) => ({
                        cells: [tableCell(`r${index}`, 21 + index * 20)],
                    })),
                ],
            };

            const { pages } = flow([mixed]);

            expect(pages.flatMap((page) => page.rows).filter((row) => row.repeated)).toEqual([]);
        });

        /**
         *  The derivation above is a GUESS, and it was wrong in one
         * direction for as long as it was the only rule: an editor's `<th>` row
         * repeated on the canvas while the `.docx` never did, because nothing
         * wrote `w:tblHeader`. A caller that knows the answer states it.
         */
        it('repeats a row that states it, whatever its cells are', () => {
            const stated: FlowTable = {
                at: 1,
                columnWidthsPx: [CELL * 10, CELL * 10],
                rows: [
                    { cells: [tableCell('HEAD', 1), tableCell('ING', 11)], repeatHeader: true },
                    ...Array.from({ length: 6 }, (_unused, index) => ({
                        cells: [tableCell(`r${index}`, 21 + index * 20)],
                    })),
                ],
            };

            const { pages } = flow([stated]);

            expect(pages.length).toBe(2);
            expect(pages[1]!.rows[0]!.repeated).toBe(true);
        });

        /**
         * And the other direction, which `||` would get wrong: header cells in
         * a row the author said does not repeat.
         */
        it('does not repeat a row of header cells that states it does not', () => {
            const declined: FlowTable = {
                at: 1,
                columnWidthsPx: [CELL * 10, CELL * 10],
                rows: [
                    {
                        cells: [tableCell('HEAD', 1, true), tableCell('ING', 11, true)],
                        repeatHeader: false,
                    },
                    ...Array.from({ length: 6 }, (_unused, index) => ({
                        cells: [tableCell(`r${index}`, 21 + index * 20)],
                    })),
                ],
            };

            const { pages } = flow([declined]);

            expect(pages.length).toBe(2);
            expect(pages.flatMap((page) => page.rows).filter((row) => row.repeated)).toEqual([]);
        });

        it('carries the caller\'s cell padding into the layout', () => {
            const padded: FlowTable = { ...grid(1), cellPaddingPx: 6 };

            const { pages } = flow([padded]);

            expect(pages[0]!.rows[0]!.heightPx).toBe(LINE + 12);
            expect(pages[0]!.rows[0]!.cells[0]!.lines[0]!.yPx).toBe(PAGE.marginTopPx + 6);
        });

        it('mixes tables and paragraphs in one flow', () => {
            const { pages, paragraphs } = flow([
                { spans: [{ text: 'before', at: 1 }], at: 1 },
                grid(2),
                { spans: [{ text: 'after', at: 200 }], at: 200 },
            ]);

            expect(paragraphs.filter(isTable).length).toBe(1);
            expect(pages[0]!.rows.length).toBe(2);
            // Text above the table and below it, with the rows between.
            expect(pages[0]!.lines.length).toBe(2);
            expect(pages[0]!.lines[1]!.yPx).toBe(PAGE.marginTopPx + LINE * 3);
        });

        it('gives a cell ONE grid column when it says nothing', () => {
            const { paragraphs } = flow([grid(1)]);
            const table = paragraphs[0];
            if (undefined === table || !isTable(table)) {
                throw new Error('expected a table');
            }

            expect(table.rows[0]!.cells[0]!.gridSpan).toBe(1);
        });

        it('honours an explicit page break before a table', () => {
            const { pages, pageStarts } = flow([
                { spans: [{ text: 'before', at: 1 }], at: 1 },
                { ...grid(1), pageBreakBefore: true },
            ]);

            expect(pages.length).toBe(2);
            expect(pages[1]!.rows.length).toBe(1);
            expect(pageStarts[0]!.rowIndex).toBe(0);
        });

        it('takes a row\'s position from its first cell\'s SPAN', () => {
            // The block's own `at` and its first span's are the same in most
            // fixtures, which hides which one is read. Here they differ.
            const table: FlowTable = {
                at: 900,
                columnWidthsPx: [CELL * 10],
                rows: [
                    { cells: [{ blocks: [{ spans: [{ text: 'r0', at: 11 }], at: 10 }] }] },
                    ...Array.from({ length: 5 }, (_unused, index) => ({
                        cells: [{ blocks: [{ spans: [{ text: `r${index + 1}`, at: 101 + index * 10 }], at: 100 + index * 10 }] }],
                    })),
                ],
            };

            const { pageStarts } = flow([table]);

            // Page two begins at row 5, whose span sits at 141 and whose block
            // starts one position earlier.
            expect(pageStarts[0]!.rowIndex).toBe(5);
            expect(pageStarts[0]!.at).toBe(141);
        });

        it('points at the LINE when text begins a page a table also appears on', () => {
            // Page two opens with the tail of a paragraph and the table follows
            // it further down. Preferring the row would attribute the page to
            // content the reader has not reached yet.
            const { pages, pageStarts } = flow([block(8, 1), grid(2)]);

            expect(pages.length).toBeGreaterThan(1);
            expect(pages[1]!.lines.length).toBeGreaterThan(0);
            expect(pages[1]!.rows.length).toBeGreaterThan(0);
            expect(pageStarts[0]!.rowIndex).toBeUndefined();
        });

        it('points at a paragraph, not a row, when text starts the page', () => {
            // A page that begins with a paragraph must not be attributed to a
            // table row further down it.
            const { pageStarts } = flow([grid(4), {
                spans: [{ text: 'x'.repeat(20) + ' ' + 'y'.repeat(20), at: 500 }],
                at: 500,
            }]);

            expect(pageStarts[0]!.rowIndex).toBeUndefined();
            expect(pageStarts[0]!.at).toBe(500 + 21);
        });
    });
});

describe('a merged cell from the editor', () => {
    const cell = (text: string, extra: Partial<FlowTableCell> = {}): FlowTableCell => ({
        blocks: [{ spans: [{ text }] }],
        ...extra,
    });

    const tableOf = (rows: FlowTableRow[], columns = 2) => {
        const { paragraphs } = paginateFlow(
            [{ rows, columnWidthsPx: Array.from({ length: columns }, () => 40) }],
            PAGE,
            { fonts: FONTS },
        );
        const [block] = paragraphs;

        return isTable(block!) ? block : undefined;
    };

    it('turns a rowSpan into a restart and its continuations', () => {
        // The editor states the merge once; the layout wants a place-holder in
        // every row it reaches, or the grid columns stop lining up.
        const table = tableOf([
            { cells: [cell('m', { rowSpan: 3 }), cell('a')] },
            { cells: [cell('b')] },
            { cells: [cell('c')] },
        ]);

        expect(table?.rows.map((row) => row.cells.map((held) => held.verticalMerge)))
            .toEqual([['restart', undefined], ['continue', undefined], ['continue', undefined]]);
    });

    it('puts the place-holder in the COLUMN the merge occupies', () => {
        // The merge is in the SECOND column here. Appending the place-holder
        // would leave the row's own cell in column two and the merge in one.
        const table = tableOf([
            { cells: [cell('a'), cell('m', { rowSpan: 2 })] },
            { cells: [cell('b')] },
        ]);

        expect(table?.rows[1]!.cells.map((held) => held.verticalMerge))
            .toEqual([undefined, 'continue']);
    });

    it('gives the place-holder the merge\'s own column span', () => {
        // A cell that is two columns wide is swallowed two columns at a time,
        // or every row below it is one column out.
        const table = tableOf([
            { cells: [cell('m', { rowSpan: 2, gridSpan: 2 }), cell('a')] },
            { cells: [cell('b')] },
        ], 3);

        expect(table?.rows[1]!.cells.map((held) => held.gridSpan)).toEqual([2, 1]);

        // The place-holder ADVANCES by the merge's own span too. A row with
        // more cells than the grid has columns left is where that shows: only
        // the one that fits is taken, and the rest have nowhere to go.
        const crowded = tableOf([
            { cells: [cell('m', { rowSpan: 2, gridSpan: 2 }), cell('a')] },
            { cells: [cell('b'), cell('c')] },
        ], 3);

        expect(crowded?.rows[1]!.cells.length).toBe(2);
        expect(crowded?.rows[1]!.cells.map((held) => held.gridSpan)).toEqual([2, 1]);
    });

    it('closes the merge after the rows it asked for', () => {
        const table = tableOf([
            { cells: [cell('m', { rowSpan: 2 }), cell('a')] },
            { cells: [cell('b')] },
            { cells: [cell('own'), cell('c')] },
        ]);

        expect(table?.rows[2]!.cells.map((held) => held.verticalMerge))
            .toEqual([undefined, undefined]);
    });

    it('leaves a table with no rowSpan exactly as it was', () => {
        const table = tableOf([
            { cells: [cell('a'), cell('b')] },
            { cells: [cell('c'), cell('d')] },
        ]);

        expect(table?.rows.map((row) => row.cells.length)).toEqual([2, 2]);
        expect(table?.rows.flatMap((row) => row.cells)
            .every((held) => undefined === held.verticalMerge)).toBe(true);
    });
});

describe('a table nested inside an editor cell', () => {
    const text = (value: string, at?: number): FlowBlock =>
        ({ spans: [{ text: value, ...(undefined === at ? {} : { at }) }] });

    const inner = (value: string, at?: number): FlowTable => ({
        rows: [{ cells: [{ blocks: [text(value, at)] }] }],
        columnWidthsPx: [40],
    });

    const outer = (blocks: FlowItem[]): FlowTable => ({
        rows: [{ cells: [{ blocks }] }],
        columnWidthsPx: [80],
    });

    const tableOf = (table: FlowTable) => {
        const [block] = paginateFlow([table], PAGE, { fonts: FONTS }).paragraphs;

        return isTable(block!) ? block : undefined;
    };

    it('lays the nested table out instead of losing it', () => {
        // A `.docx` cell may hold a table and an editor's may too — a grid
        // inside a grid is how a page layout is built in one.
        const table = tableOf(outer([inner('deep')]));
        const held = table?.rows[0]!.cells[0]!.paragraphs[0];

        expect(undefined === held ? false : isTable(held)).toBe(true);
    });

    it('keeps the blocks around it in document order', () => {
        const table = tableOf(outer([text('above'), inner('deep'), text('below')]));
        const held = table?.rows[0]!.cells[0]!.paragraphs ?? [];

        expect(held.map(isTable)).toEqual([false, true, false]);
    });

    it('nests as far down as the editor writes it', () => {
        const twice = outer([{
            rows: [{ cells: [{ blocks: [inner('deep')] }] }],
            columnWidthsPx: [60],
        }]);
        const first = tableOf(twice)?.rows[0]!.cells[0]!.paragraphs[0];
        const second = undefined !== first && isTable(first)
            ? first.rows[0]!.cells[0]!.paragraphs[0]
            : undefined;

        expect(undefined === second ? false : isTable(second)).toBe(true);
    });

    it('finds a row position inside a nested table', () => {
        // The position wanted is the EARLIEST in the row, however deep it is
        // written — an editor pointing a caret at the row needs the innermost
        // handle, not the table's own.
        const { pageStarts } = paginateFlow(
            [
                // Exactly a page of text — the column holds two of these
                // words to a line — so the next page BEGINS at the table
                // rather than part-way through the paragraph.
                { spans: Array.from({ length: 10 }, () => ({ text: 'xxxxxxxxxx ' })) },
                outer([inner('deep', 42)]),
            ],
            PAGE,
            { fonts: FONTS },
        );
        const started = pageStarts.find((start) => undefined !== start.rowIndex);

        expect(started?.at).toBe(42);
    });

    it('looks PAST a nested table that carries no position', () => {
        // A table the editor gave no handle to answers nothing; the search has
        // to go on to the block after it rather than give up on the row.
        const { pageStarts } = paginateFlow(
            [
                { spans: Array.from({ length: 10 }, () => ({ text: 'xxxxxxxxxx ' })) },
                outer([inner('deep'), text('after', 7)]),
            ],
            PAGE,
            { fonts: FONTS },
        );
        const started = pageStarts.find((start) => undefined !== start.rowIndex);

        expect(started?.at).toBe(7);
    });
});

describe('a block that boxes itself', () => {
    const side: BorderSide = {
        widthPx: 4, style: 'single' as BorderStyle, colorHex: '#000000', spacePx: 6,
    };
    const BOX: BoxBorders = { top: side, left: side, bottom: side, right: side };

    /** `count` blocks, the middle one boxed unless `plain`. */
    const blocks = (count: number, plain: boolean): FlowItem[] =>
        Array.from({ length: count }, (_, index) => ({
            spans: [{ text: `L${index}`, at: index * 10 }],
            at: index * 10,
            ...(1 === index && !plain ? { style: { borders: BOX } } : {}),
        }));

    it('carries the border through to the paragraph it lays out', () => {
        // The editor draws the box in its own CSS; what the ENGINE needs it
        // for is the room, and it cannot take room it was never told about.
        const boxed = paragraphAt(flow(blocks(3, false)), 1);

        expect(boxed.style.borders).toBe(BOX);
        expect(paragraphAt(flow(blocks(3, true)), 1).style.borders).toBeUndefined();
    });

    it('takes the room out of the page, so the BREAK moves', () => {
        // Measured against LibreOffice: a bordered paragraph steps 12.5pt
        // where a plain one steps 11.5, and more again with `w:space` — here
        // ten points either side, which is a whole line of this page. Five
        // plain blocks fit; with one of them boxed, the fifth is pushed over.
        expect(flow(blocks(5, true)).pages).toHaveLength(1);
        expect(flow(blocks(5, false)).pages.length).toBeGreaterThan(1);
    });

    it('leaves an unboxed run of blocks paginating exactly as before', () => {
        const plain = flow(blocks(4, true));

        expect(plain.pages).toHaveLength(1);
        expect(plain.paragraphs).toHaveLength(4);
    });
});

describe('the seam the editor paginates through', () => {
    /**
     * Every `ParagraphStyle` key, said to be one of two things.
     *
     * `borders` was missing from `BlockStyle` for the whole life of the flow
     * seam, so a boxed paragraph paginated as a plain one and the
     * editor's page breaks drifted from the exported document's. Nothing
     * caught it because nothing was watching the DIFFERENCE between what the
     * layout uses and what the editor can say.
     *
     * This is that watch, and it is the compiler's rather than a reviewer's:
     * add a key to `ParagraphStyle` and this object stops type-checking until
     * someone decides which kind it is. `carried` means the editor supplies it
     * and a break depends on it. `drawn by the editor` means the editor paints
     * it in its own CSS and the engine's copy changes no break.
     */
    const SEAM: Record<
        keyof ParagraphStyle,
        'carried' | 'drawn by the editor' | 'read from the file only'
    > = {
        // These decide where a break falls, so the editor must be able to say them.
        spaceBeforePx: 'carried',
        spaceAfterPx: 'carried',
        lineHeightPx: 'carried',
        indentLeftPx: 'carried',
        indentRightPx: 'carried',
        indentFirstLinePx: 'carried',
        pageBreakBefore: 'carried',
        keepLinesTogether: 'carried',
        keepWithNext: 'carried',
        widowControl: 'carried',
        borders: 'carried',

        // `w:suppressLineNumbers`, which decides no break at all: it takes a
        // paragraph out of `w:lnNumType`'s count, and the editor
        // neither numbers lines nor has a way to say this.
        suppressLineNumbers: 'read from the file only',

        // A section's `w:docGrid`, which DOES decide where breaks fall — under
        // an 18pt grid a 1.5-spaced paragraph steps 27.00 rather than 17.25 —
        // and which the editor has no way to say: a ProseMirror
        // document is not ruled to a pitch. So it is neither of the two the
        // seam had, and saying so is the point: a field the editor
        // cannot fill is worse invented than left where it is.
        gridPitchPx: 'read from the file only',
        // Which KIND of gridded height this is — a floor or a multiple. Read
        // with the pitch and belonging to it; a ProseMirror paragraph has
        // neither to give.
        gridFloor: 'read from the file only',

        // These move ink WITHIN a line the editor has already drawn itself.
        alignment: 'drawn by the editor',
        lineRule: 'drawn by the editor',
        tabStops: 'drawn by the editor',
        decimalSymbol: 'drawn by the editor',
        defaultTabPx: 'drawn by the editor',
    };

    /**
     * The same watch over a TABLE, which the seam also carries.
     *
     * `borders` was missing here for the same reason it was missing from
     * `BlockStyle`: it looks like ink. It is room — every horizontal rule sits
     * in a gap of its own width — so a twenty-row table
     * paginated without it drifts by nearly two lines.
     */
    const TABLE_SEAM: Record<keyof Table, 'carried' | 'drawn by the editor' | 'read from the file only'> = {
        rows: 'carried',
        columnWidthsPx: 'carried',
        spaceBeforePx: 'carried',
        spaceAfterPx: 'carried',
        pageBreakBefore: 'carried',
        // CARRIED, but NOT to be combined with a padding that already includes
        // them. The admin editor measures `cellBoxPx` off a real cell — "padding
        // PLUS borders" — and passes it as `cellPaddingPx`, so wiring `borders`
        // through from the same CSS would count them twice. It does not today;
        // this is here because the next person to look will be tempted.
        //
        // The two accountings are not interchangeable either. Padding is per
        // CELL, so N rows pay 2N of it; a rule sits in a gap of its own width
        // BETWEEN contents, so N rows pay N + 1. They agree only at
        // one row. Which is right depends on the reference: the browser is the
        // editor's, LibreOffice is the export's, and where they disagree the
        // export is the one a reader ends up holding.
        borders: 'carried',
        // One uniform `cellPaddingPx` becomes all four, which is what a
        // ProseMirror table has to give.
        cellMarginLeftPx: 'carried',
        cellMarginRightPx: 'carried',
        cellMarginTopPx: 'carried',
        cellMarginBottomPx: 'carried',

        shadingFill: 'drawn by the editor',
        // The editor's columns are already resolved to pixels by the time they
        // reach the seam, so nothing here re-sizes the table.
        preferredWidthFraction: 'read from the file only',
        preferredWidthPx: 'read from the file only',
        alignment: 'read from the file only',
        indentPx: 'read from the file only',
    };

    const CELL_SEAM: Record<keyof TableCell, 'carried' | 'drawn by the editor' | 'read from the file only'> = {
        paragraphs: 'carried',
        gridSpan: 'carried',
        verticalMerge: 'carried',
        borders: 'drawn by the editor',
        shadingFill: 'drawn by the editor',
        verticalAlign: 'drawn by the editor',
        // `w:tcMar` narrows a cell's text column, so it WOULD move a break —
        // but a ProseMirror cell has no per-cell padding to say it with, and
        // inventing one the editor cannot fill is worse than the gap. Recorded
        // rather than carried; the uniform `cellPaddingPx` covers the case the
        // editor actually has.
        margins: 'read from the file only',
        // Read so the model is whole, laid out by nobody yet,
        // and a ProseMirror cell cannot be turned on its side anyway.
        textDirection: 'read from the file only',
    };

    /**
     * And the third seam: a SPAN, which becomes a run on a line.
     *
     * `TextSpan` carries six of `StyledRun`'s fifteen. Most of the rest is ink
     * the editor paints itself — but three of them move a BREAK, and those are
     * the ones worth saying out loud rather than leaving to be rediscovered a
     * fourth time.
     */
    const SPAN_SEAM: Record<keyof StyledRun, 'carried' | 'drawn by the editor' | 'read from the file only'> = {
        text: 'carried',
        font: 'carried',
        sizePx: 'carried',

        colorHex: 'drawn by the editor',
        highlightHex: 'drawn by the editor',
        underline: 'drawn by the editor',
        strike: 'drawn by the editor',

        // These three DO move a break, and none of them is sayable today.
        //
        // `letterSpacingPx` changes an advance, so it changes where a line
        // wraps: an editor tracking a heading would paginate wide. A run
        // `border` keeps `space + width` clear on all four sides.
        // `baselineShiftPx` grows the line by the shift — though the
        // editor's superscript is a smaller `fontSizePx`, which it CAN say,
        // and at 0.58 of the size the shift stays inside the line anyway.
        //
        // Not added: a ProseMirror span has none of them to give, and a field
        // the editor cannot fill is worse than the gap it papers over. Recorded
        // here so the next person meets the decision rather than the silence.
        letterSpacingPx: 'read from the file only',
        border: 'read from the file only',
        baselineShiftPx: 'read from the file only',

        // A fourth that moves a break, and the widest of them: kerning `AV`
        // pairs is worth 11.60pt over ten characters. A ProseMirror
        // span cannot say `w:kern` either — but unlike the three above, the
        // ABSENT case is the common one and is what this engine now does, so
        // an editor span silently gets the right answer for once.
        kerned: 'read from the file only',

        // A note belongs to the document, not to a span the editor is holding.
        footnoteId: 'read from the file only',
        field: 'read from the file only',
        fieldFormat: 'read from the file only',
        // The editor measures its own opaque nodes and gives the engine a
        // height, so it never hands over a picture to size.
        image: 'read from the file only',
        shape: 'read from the file only',
        // A `w:ruby` gloss makes its line taller by a whole second line and
        // can decide the run's advance, so it moves a break twice
        // over — and a ProseMirror span has no gloss to hand across.
        ruby: 'read from the file only',
    };

    it('says what a SPAN can carry, and what it cannot', () => {
        expect(SPAN_SEAM.text).toBe('carried');
        expect(SPAN_SEAM.letterSpacingPx).toBe('read from the file only');
    });

    it('classifies every table key, and carries what a break depends on', () => {
        expect(TABLE_SEAM.borders).toBe('carried');
        expect(CELL_SEAM.paragraphs).toBe('carried');
    });

    it('keeps room for a table’s rules, so the BREAK moves', () => {
        // Five one-line rows fit this page. Give the table one-point rules and
        // it needs six gaps of its own width on top — enough to push the last
        // row over, which is the whole of what the editor has to agree with.
        const side: BorderSide = {
            widthPx: LINE / 2, style: 'single' as BorderStyle, colorHex: '#000000',
        };
        const rows = Array.from({ length: 4 }, (_, index) => ({
            cells: [tableCell(`R${index}`, index * 10)],
        }));
        const plain: FlowTable = { rows, columnWidthsPx: [CELL * 10], at: 0 };

        expect(flow([plain]).pages).toHaveLength(1);
        expect(flow([{ ...plain, borders: { top: side, bottom: side, insideH: side } }])
            .pages.length).toBeGreaterThan(1);
    });

    it('lets the editor say everything a BREAK depends on', () => {
        // `pageBreakBefore` rides on the block rather than its style, which is
        // why it is checked apart from the rest.
        const carried = Object.entries(SEAM)
            .filter(([, kind]) => 'carried' === kind)
            .map(([key]) => key)
            .filter((key) => 'pageBreakBefore' !== key);
        const sayable = new Set(Object.keys({
            fontFamily: 0, fontSizePx: 0, bold: 0, italic: 0, lineHeightPx: 0,
            spaceBeforePx: 0, spaceAfterPx: 0, indentLeftPx: 0, indentRightPx: 0,
            indentFirstLinePx: 0, keepLinesTogether: 0, keepWithNext: 0,
            widowControl: 0, borders: 0,
        } satisfies Record<keyof BlockStyle, number>));

        expect(carried.filter((key) => !sayable.has(key))).toEqual([]);
    });

    it('leaves a block’s own pageBreakBefore reaching the layout', () => {
        const pages = flow([
            { spans: [{ text: 'one', at: 0 }], at: 0 },
            { spans: [{ text: 'two', at: 10 }], at: 10, pageBreakBefore: true },
        ]).pages;

        expect(pages).toHaveLength(2);
    });
});
