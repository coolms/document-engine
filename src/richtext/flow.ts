import type { BoxBorders } from '../layout/borders.js';
import type { StyledRun } from '../layout/line-breaker.js';
import {
    layoutPages,
    type Block,
    type Page,
    type PageGeometry,
    type Paragraph,
    type ParagraphStyle,
    type PlacedRow,
    type Table,
    type TableCell,
} from '../layout/page-layout.js';
import type { FontCatalogue } from '../word/font-catalogue.js';

/**
 * Paginating an EDITOR's document — the shape a rich-text editor already has,
 * rather than the shape a `.docx` has.
 *
 * ## Why the caller supplies the coordinates
 *
 * An editor needs page boundaries back as positions in ITS OWN document model,
 * so it can put a spacer there or tell the user which page the caret is on.
 * ProseMirror's position arithmetic has rules of its own — a text node counts
 * its UTF-16 length, a leaf counts one, a container counts two plus its content
 * — and re-deriving them here from a JSON tree would be a second implementation
 * that agrees with the first until it does not.
 *
 * So each span carries an `at` handle that means whatever the caller wants, and
 * this module does arithmetic on it without interpreting it. The engine stays
 * framework-agnostic and the editor keeps the only copy of its own rules.
 */

export interface TextSpan {
    readonly text: string;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly fontFamily?: string;
    readonly fontSizePx?: number;
    /**
     * The caller's position for this span's first character. Offsets within the
     * span are added to it, so it must advance one per UTF-16 unit — which is
     * exactly how a ProseMirror text position behaves.
     */
    readonly at?: number;
}

/** Per-block overrides; anything absent comes from the flow's base style. */
export interface BlockStyle {
    readonly fontFamily?: string;
    readonly fontSizePx?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly lineHeightPx?: number;
    readonly spaceBeforePx?: number;
    readonly spaceAfterPx?: number;
    readonly indentLeftPx?: number;
    readonly indentRightPx?: number;
    readonly indentFirstLinePx?: number;
    readonly keepLinesTogether?: boolean;
    readonly keepWithNext?: boolean;
    readonly widowControl?: boolean;
    /**
     * `w:pBdr` — a box round the block, which costs the FLOW room.
     *
     * Here because a border is not only ink: measured against LibreOffice,
     * a bordered paragraph steps 12.5pt where a plain one steps 11.5, and
     * more again with `w:space`. An editor that paginates without it puts
     * its page breaks in a different place from the document it exports —
     * which is the whole of what a paged editor is for. The editor draws
     * the box itself, in its own CSS; this is what the BREAKS need.
     */
    readonly borders?: BoxBorders;
}

export interface FlowBlock {
    readonly spans: readonly TextSpan[];
    readonly style?: BlockStyle;
    /** An explicit break the author placed before this block. */
    readonly pageBreakBefore?: boolean;
    /** The caller's position for the block's first character. */
    readonly at?: number;
}

/**
 * A table in the editor's document.
 *
 * The caller supplies the column widths because only it knows them: a
 * ProseMirror table carries `colwidth` per cell, and where that is absent the
 * columns divide the writing width evenly — a rule about the EDITOR's CSS,
 * not about documents.
 */
export interface FlowTableCell {
    /**
     * What the cell holds — paragraphs, and tables nested inside it.
     *
     * A `FlowItem` rather than a block, because a cell in a `.docx` may hold a
     * table and an editor's may too: a grid inside a grid is how a page layout
     * is built in one. Named `blocks` still, since the type says what is in it.
     */
    readonly blocks: readonly FlowItem[];
    readonly gridSpan?: number;
    /**
     * How many rows the cell covers — a ProseMirror `rowspan`.
     *
     * An editor states a merge ONCE, on the cell that opens it, and the rows
     * beneath simply have one cell fewer. The layout wants the opposite: a
     * place-holder in every row the merge reaches, so the grid columns still
     * line up. Turning one into the other belongs here, at the seam between
     * the editor's shape and the engine's.
     */
    readonly rowSpan?: number;
    /** A header CELL. A row of them is a header row, which repeats. */
    readonly isHeader?: boolean;
}

export interface FlowTableRow {
    readonly cells: readonly FlowTableCell[];
    /**
     * `w:trHeight` — a height the AUTHOR set, in px.
     *
     * Absent means the row is as tall as its tallest cell, which is what a
     * table does unasked and what every row did before this existed. Stated, it
     * is a FLOOR: the row grows past it for content that needs more, matching
     * the `atLeast` rule Word writes by default. The layout already understood
     * both (see `TableRow.heightPx`); until now nothing in an editor's document
     * could reach it.
     */
    readonly heightPx?: number;
    /**
     * `w:tblHeader` — "repeat this row at the top of every page", stated.
     *
     * ⚠️ Absent falls back to the old rule: a row of header CELLS repeats. That
     * rule is a guess, and it was wrong in one direction for as long as it was
     * the only one — the canvas repeated a `<th>` row that the `.docx` never
     * did, because nothing wrote `w:tblHeader`. A document whose editor states
     * the fact says so here instead, and then the sheet on screen and the
     * printed page break in the same place.
     */
    readonly repeatHeader?: boolean;
}

export interface FlowTable {
    readonly rows: readonly FlowTableRow[];
    readonly columnWidthsPx: readonly number[];
    /** Cell padding, which in an editor is one symmetric number from the CSS. */
    readonly cellPaddingPx?: number;
    /**
     * The rules the table draws, which are ROOM before they are ink.
     *
     * Measured against LibreOffice: every horizontal rule sits
     * in a gap of its own width, so a table of N rows is N + 1 rule-widths
     * taller than its text. The editor paints them in CSS; without this
     * the engine keeps no gap and a twenty-row table's breaks drift by
     * nearly two lines from the document it exports.
     */
    readonly borders?: BoxBorders;
    /**
     * The table's OWN box above and below its rows.
     *
     * A grid layout in an editor is a table wrapped in a div with padding, a
     * border and margins; without somewhere to put that, the rows are right
     * and the thing they sit in is invisible to the layout.
     */
    readonly spaceBeforePx?: number;
    readonly spaceAfterPx?: number;
    readonly pageBreakBefore?: boolean;
    readonly at?: number;
}

/** What a flow is made of. A table has rows; a paragraph has spans. */
export type FlowItem = FlowBlock | FlowTable;

export function isFlowTable(item: FlowItem): item is FlowTable {
    return 'rows' in item;
}

export interface FlowStyle {
    readonly fontFamily: string;
    readonly fontSizePx: number;
    readonly lineHeightPx?: number;
    readonly spaceBeforePx: number;
    readonly spaceAfterPx: number;
    readonly indentLeftPx: number;
    readonly indentRightPx: number;
    readonly indentFirstLinePx: number;
    readonly keepLinesTogether: boolean;
    readonly keepWithNext: boolean;
}

export const DEFAULT_FLOW_STYLE: FlowStyle = {
    fontFamily: 'Calibri',
    fontSizePx: 16,
    spaceBeforePx: 0,
    spaceAfterPx: 0,
    indentLeftPx: 0,
    indentRightPx: 0,
    indentFirstLinePx: 0,
    keepLinesTogether: false,
    keepWithNext: false,
};

export interface FlowPageStart {
    /** Index of the page that BEGINS here. Page 0 begins the document and is never listed. */
    readonly page: number;
    readonly blockIndex: number;
    /** Offset into the block's concatenated text, in UTF-16 units. */
    readonly textOffset: number;
    /** The caller's coordinate, or null when the spans carried none. */
    readonly at: number | null;
    /**
     * True when the page starts PART-WAY through a block.
     *
     * The two cases need different treatment in an editor: between blocks a
     * spacer can go before the next block, while inside one the block itself has
     * to be split visually.
     */
    readonly insideBlock: boolean;
    /** Set when the page begins at a table ROW rather than at text. */
    readonly rowIndex?: number;
}

export interface FlowPagination {
    readonly pages: Page[];
    /** The layout blocks the flow was turned into — paragraphs and tables. */
    readonly paragraphs: Block[];
    /** Where each page after the first begins. Its length is `pages.length - 1`. */
    readonly pageStarts: readonly FlowPageStart[];
}

export interface FlowOptions {
    readonly fonts: FontCatalogue;
    readonly base?: Partial<FlowStyle>;
    readonly widowOrphanControl?: boolean;
}

export function paginateFlow(
    items: readonly FlowItem[],
    geometry: PageGeometry,
    options: FlowOptions,
): FlowPagination {
    const base: FlowStyle = { ...DEFAULT_FLOW_STYLE, ...options.base };
    const paragraphs: Block[] = items.map((item) => isFlowTable(item)
        ? toTable(item, base, options.fonts)
        : toParagraph(item, base, options.fonts));

    const pages = layoutPages(
        paragraphs,
        geometry,
        undefined === options.widowOrphanControl ? {} : { widowOrphanControl: options.widowOrphanControl },
    );

    const pageStarts: FlowPageStart[] = [];
    for (let index = 1; index < pages.length; index++) {
        const page = pages[index]!;
        const firstLine = page.lines[0];
        // A repeated header is drawn again at the top; the page does not BEGIN
        // there, and pointing an editor at it would send the caret backwards.
        const firstRow = page.rows.find((row) => !row.repeated);

        const rowIsFirst = undefined !== firstRow
            && (undefined === firstLine || firstRow.yPx <= firstLine.yPx);

        if (rowIsFirst) {
            const table = items[firstRow.blockIndex];
            pageStarts.push({
                page: index,
                blockIndex: firstRow.blockIndex,
                textOffset: 0,
                rowIndex: firstRow.rowIndex,
                at: undefined === table || !isFlowTable(table)
                    ? null
                    : positionOfRow(table, firstRow.rowIndex),
                insideBlock: firstRow.rowIndex > 0,
            });

            continue;
        }

        if (undefined === firstLine) {
            // A page with nothing on it cannot be pointed at. It should not
            // happen, and a start of "nowhere" would be worse than omitting it.
            continue;
        }

        const block = items[firstLine.paragraphIndex];
        pageStarts.push({
            page: index,
            blockIndex: firstLine.paragraphIndex,
            textOffset: firstLine.line.startsAt,
            at: undefined === block || isFlowTable(block)
                ? null
                : positionAt(block, firstLine.line.startsAt),
            insideBlock: firstLine.lineIndex > 0,
        });
    }

    return { pages, paragraphs, pageStarts };
}

/**
 * A flow table as the layout's table.
 *
 * A row counts as a header only when EVERY cell in it is one — which is how a
 * rich-text editor models a header row, and stops a single header cell in an
 * ordinary row from making that row repeat on every page.
 */
function toTable(table: FlowTable, base: FlowStyle, fonts: FontCatalogue): Table {
    const padding = table.cellPaddingPx ?? 0;

    // What each row still owes to a merge opened above it, by grid COLUMN.
    const open = new Map<number, { rows: number; gridSpan: number }>();

    const rows = table.rows.map((row) => {
        const cells: TableCell[] = [];
        const own = [...row.cells];
        let column = 0;

        // Walked by column rather than by cell: a row under a merge has one
        // cell fewer than the grid has columns, and appending the place-holder
        // would put it at the END of the row instead of in the column the
        // merge occupies — which is the column everything below lines up on.
        while (column < table.columnWidthsPx.length) {
            const held = open.get(column);
            if (undefined !== held) {
                cells.push({ paragraphs: [], gridSpan: held.gridSpan, verticalMerge: 'continue' });
                held.rows -= 1;
                if (held.rows <= 0) {
                    open.delete(column);
                }
                column += held.gridSpan;

                continue;
            }

            const cell = own.shift();
            if (undefined === cell) {
                break;
            }

            const gridSpan = Math.max(1, cell.gridSpan ?? 1);
            const rowSpan = Math.max(1, cell.rowSpan ?? 1);
            cells.push({
                paragraphs: cell.blocks.map((block) => isFlowTable(block)
                    ? toTable(block, base, fonts)
                    : toParagraph(block, base, fonts)),
                gridSpan,
                ...(rowSpan > 1 ? { verticalMerge: 'restart' as const } : {}),
            });
            if (rowSpan > 1) {
                open.set(column, { rows: rowSpan - 1, gridSpan });
            }
            column += gridSpan;
        }

        return {
            cells,
            // ⚠️ The STATED fact first. `??` and not `||`, because an explicit
            // `false` is an author saying "these are header cells and this row
            // does not repeat" — which is exactly the case the fallback below
            // gets wrong.
            isHeader: row.repeatHeader
                ?? (0 < row.cells.length && row.cells.every((cell) => true === cell.isHeader)),
            cantSplit: false,
            // A floor, not a fixed height: `heightRule` is left absent, which
            // the layout reads as `atLeast`. An author who sets 40px on a row
            // whose text needs 60 gets 60 — the same answer Word gives, and the
            // only one that cannot hide their content.
            ...(undefined === row.heightPx ? {} : { heightPx: row.heightPx }),
        };
    });

    return {
        rows,
        columnWidthsPx: table.columnWidthsPx,
        ...(undefined === table.borders ? {} : { borders: table.borders }),
        cellMarginLeftPx: padding,
        cellMarginRightPx: padding,
        cellMarginTopPx: padding,
        cellMarginBottomPx: padding,
        spaceBeforePx: table.spaceBeforePx ?? 0,
        spaceAfterPx: table.spaceAfterPx ?? 0,
        pageBreakBefore: true === table.pageBreakBefore,
    };
}

/** Where a table row begins, in the caller's coordinates. */
function positionOfRow(table: FlowTable, rowIndex: number): number | null {
    const row = table.rows[rowIndex];
    if (undefined === row) {
        return null;
    }

    // The first thing in the row that has a position: its first cell's first
    // block, or that block's first span. A nested table is asked the same
    // question of its own first row, since the position wanted is the earliest
    // in the row however deep it is written.
    for (const cell of row.cells) {
        for (const block of cell.blocks) {
            if (isFlowTable(block)) {
                const inner = positionOfRow(block, 0);
                if (null !== inner) {
                    return inner;
                }

                continue;
            }

            const span = block.spans[0];
            if (undefined !== span?.at) {
                return span.at;
            }
            if (undefined !== block.at) {
                return block.at;
            }
        }
    }

    return table.at ?? null;
}

function toParagraph(block: FlowBlock, base: FlowStyle, fonts: FontCatalogue): Paragraph {
    const style = block.style ?? {};
    const family = style.fontFamily ?? base.fontFamily;
    const sizePx = style.fontSizePx ?? base.fontSizePx;

    const runs: StyledRun[] = block.spans.map((span) => ({
        text: span.text,
        sizePx: span.fontSizePx ?? sizePx,
        font: fonts.resolve(
            span.fontFamily ?? family,
            span.bold ?? style.bold ?? false,
            span.italic ?? style.italic ?? false,
        ).font,
    }));

    // An empty block still occupies a line, and its height comes from the
    // formatting it would have had. With no run the line is zero high and the
    // page holds more than it should.
    if (0 === runs.length) {
        runs.push({
            text: '',
            sizePx,
            font: fonts.resolve(family, style.bold ?? false, style.italic ?? false).font,
        });
    }

    const lineHeightPx = style.lineHeightPx ?? base.lineHeightPx;
    const paragraphStyle: ParagraphStyle = {
        spaceBeforePx: style.spaceBeforePx ?? base.spaceBeforePx,
        spaceAfterPx: style.spaceAfterPx ?? base.spaceAfterPx,
        indentLeftPx: style.indentLeftPx ?? base.indentLeftPx,
        indentRightPx: style.indentRightPx ?? base.indentRightPx,
        indentFirstLinePx: style.indentFirstLinePx ?? base.indentFirstLinePx,
        pageBreakBefore: true === block.pageBreakBefore,
        keepLinesTogether: style.keepLinesTogether ?? base.keepLinesTogether,
        keepWithNext: style.keepWithNext ?? base.keepWithNext,
        ...(undefined === lineHeightPx ? {} : { lineHeightPx }),
        ...(undefined === style.widowControl ? {} : { widowControl: style.widowControl }),
        ...(undefined === style.borders ? {} : { borders: style.borders }),
    };

    return { runs, style: paragraphStyle };
}

/**
 * The caller's position for a text offset within a block.
 *
 * Resolved through the SPANS rather than from the block's own `at` plus the
 * offset: inline content the editor counts but this engine never sees — an
 * image, a field chip — sits between spans and occupies a position without
 * occupying any text. Adding the offset to the block's start would drift past
 * every one of them.
 */
function positionAt(block: FlowBlock, textOffset: number): number | null {
    let consumed = 0;

    for (const span of block.spans) {
        if (undefined === span.at) {
            return null;
        }
        if (textOffset < consumed + span.text.length) {
            return span.at + (textOffset - consumed);
        }
        consumed += span.text.length;
    }

    // Exactly at the end: the last span's position plus its whole length. With
    // no spans at all there is only the block's own start to offer.
    const last = block.spans[block.spans.length - 1];
    if (undefined === last || undefined === last.at) {
        return block.at ?? null;
    }

    return last.at + last.text.length;
}
