import { alignLine } from './alignment.js';
import {
    borderRoom,
    borderStandoff,
    resolveCellBorders,
    sameBorders,
    strongerBorder,
    type BorderSide,
    type BoxBorders,
    type PlacedParagraphBorder,
} from './borders.js';
import type { Line } from './line-breaker.js';
// Types only, so this never becomes a runtime import cycle with page-layout.
import type { Paragraph, ParagraphStyle, PlacedLine } from './page-layout.js';

/**
 * Tables: the second thing a page is made of.
 *
 * A table is not a paragraph with lines in it. Its height comes from the
 * TALLEST cell in each row, its cells are columns of their own with their own
 * line breaking, and it breaks across pages between rows rather than between
 * lines. All of that is separate enough to live here.
 */

/**
 * `w:vAlign` — where a cell's content sits when the row is TALLER than it.
 *
 * Invisible in a row that fits its own text, which is why it only ever shows
 * beside a taller neighbour or under a `w:trHeight`.
 */
export type CellVerticalAlign = 'top' | 'center' | 'bottom';

/**
 * `w:hRule` — how to read the height a row asks for.
 *
 * `exact` is the only one that can make a row SMALLER than its text. Word's
 * `auto` is not listed: measured against LibreOffice, a row that states a
 * height and no rule keeps that height as a minimum, which is `atLeast`.
 */
export type RowHeightRule = 'atLeast' | 'exact';

/**
 * `w:vMerge` — a cell joined to the one above it in the same grid column.
 *
 * `restart` opens a merge and holds the content; `continue` is swallowed by it.
 * A swallowed cell keeps its place in the row so the grid columns still line
 * up, but it draws nothing and its own text is ignored — which is Word's rule,
 * confirmed against LibreOffice, where text in a continued cell vanished.
 */
export type VerticalMerge = 'restart' | 'continue';

/**
 * What a document is made of. A paragraph has runs; a table has rows, and that
 * is how the two are told apart — a discriminator field would have to be added
 * to every paragraph ever constructed to say something the shape already says.
 *
 * Here rather than beside the page flow because a CELL holds blocks too, and
 * the check has to be callable from the code that stacks them without a runtime
 * import back the other way.
 */
export type Block = Paragraph | Table;

export function isTable(block: Block): block is Table {
    return 'rows' in block;
}

/**
 * `w:tcMar` — one cell's padding, each side standing in for the table's.
 *
 * Side by side rather than all four at once: `w:tcMar` may declare any subset,
 * and the sides it leaves out keep the table's `w:tblCellMar`.
 */
export interface CellMargins {
    readonly leftPx?: number;
    readonly rightPx?: number;
    readonly topPx?: number;
    readonly bottomPx?: number;
}

/** Every side resolved, the table's standing in wherever the cell was silent. */
export interface ResolvedCellMargins {
    readonly leftPx: number;
    readonly rightPx: number;
    readonly topPx: number;
    readonly bottomPx: number;
}

export interface TableCell {
    /**
     * What the cell holds: paragraphs, and tables nested inside it.
     *
     * Named `paragraphs` for the same reason `WordDocument.paragraphs` is —
     * the type says what is in it, and renaming would touch every caller to
     * say nothing new.
     */
    readonly paragraphs: readonly Block[];
    /** How many grid columns this cell covers — `w:gridSpan`. */
    readonly gridSpan: number;
    /** `w:tcBorders`. Every side declared here beats the table's. */
    readonly borders?: BoxBorders;
    /** `w:tcPr/w:shd/@w:fill` as `#RRGGBB`, when it is not `auto`. */
    readonly shadingFill?: string;
    /** Absent means top, which is what a cell without `w:vAlign` gets. */
    readonly verticalAlign?: CellVerticalAlign;
    /** Absent means an ordinary cell, standing in one row. */
    readonly verticalMerge?: VerticalMerge;
    /**
     * `w:textDirection` — this cell's text turned a quarter turn.
     *
     * Read and carried; NOT yet laid out or drawn. The measurement that
     * decides the layout is this: a turned line's length is the ROW's height,
     * and that height comes from the cells that are NOT turned — so it is a
     * second measuring pass, not a circle. A cell turned with no neighbour to
     * set the height has no answer to copy: LibreOffice drops such a table
     * entirely, so this engine leaves it upright, which is what it does today.
     */
    readonly textDirection?: CellTextDirection;

    /**
     * `w:tcMar`, overriding the table's `w:tblCellMar` side by side.
     *
     * Measured against LibreOffice: 500 twips of left put the text 25pt in
     * from the cell's edge, 800 of right moved the wrap in by 40, and top and
     * bottom each made the ROW that much taller with the text still at the top
     * of it. A neighbouring cell was untouched by any of them.
     */
    readonly margins?: CellMargins;
}

export interface TableRow {
    readonly cells: readonly TableCell[];
    /**
     * `w:tblHeader` — repeat this row at the top of every page the table
     * continues onto. Only leading rows can be headers; a header in the middle
     * of a table is not one.
     */
    readonly isHeader: boolean;
    /**
     * `w:cantSplit` — this row moves whole rather than breaking across a page.
     *
     * Measured: a 40-line row so marked left the page it could not
     * fit entirely empty of the table, header and all. A row that fits on NO
     * page is split regardless, because the flag then asks for something no
     * page can give.
     */
    readonly cantSplit: boolean;
    /** `w:trHeight`. Absent means the row is as tall as its tallest cell. */
    readonly heightPx?: number;
    /** Ignored without {@link heightPx}. Absent means `atLeast`. */
    readonly heightRule?: RowHeightRule;
}

/**
 * `w:textDirection` — a cell whose text is turned on its side.
 *
 * `btLr` reads bottom-to-top and `tbRl` top-to-bottom; measured off the
 * PDF's text matrix, they are `[0 1 -1 0]` and `[0 -1 1 0]` — a quarter
 * turn anticlockwise and clockwise. Absent means upright, which is every
 * cell this engine has drawn until now.
 */
export type CellTextDirection = 'btLr' | 'tbRl';

/** Where a table sits in the column it lands in — `w:tblPr/w:jc`. */
export type TableAlignment = 'left' | 'center' | 'right';

export interface Table {
    readonly rows: readonly TableRow[];
    /** `w:tblBorders`, including the `insideH`/`insideV` a cell cannot declare. */
    readonly borders?: BoxBorders;
    /** `w:tblPr/w:shd/@w:fill`, which a cell's own shading covers over. */
    readonly shadingFill?: string;
    /** One width per GRID column, from `w:tblGrid`. */
    readonly columnWidthsPx: readonly number[];
    /**
     * `w:tblW` as a fraction of the column the table lands in.
     *
     * Carried unresolved because the reader does not know what it will land in
     * — the same table is one width on the page and another inside a cell.
     * Measured against LibreOffice: fifty percent of an A4 text column came
     * out at 269.29pt, which is half of 538.58 exactly.
     */
    readonly preferredWidthFraction?: number;
    /** `w:tblW` stated in twips, which needs nothing to resolve against. */
    readonly preferredWidthPx?: number;
    /** `w:tblPr/w:jc`, absent meaning left. */
    readonly alignment?: TableAlignment;
    /**
     * `w:tblInd`, which counts to the CELL'S TEXT and not the table's edge.
     *
     * So the table's own left edge lands a cell margin further LEFT than the
     * number says — measured off LibreOffice, where half an inch of indent
     * moved a table's border 31.1pt and its text the full 36. Present but
     * unresolvable (a percentage) is nought rather than absent: LibreOffice
     * still pulled the table left by the cell margin.
     */
    readonly indentPx?: number;
    readonly cellMarginLeftPx: number;
    readonly cellMarginRightPx: number;
    readonly cellMarginTopPx: number;
    readonly cellMarginBottomPx: number;
    readonly spaceBeforePx: number;
    readonly spaceAfterPx: number;
    readonly pageBreakBefore: boolean;
}

export interface PlacedCell {
    /** Left edge of the CELL, margins not yet applied. */
    readonly xPx: number;
    readonly widthPx: number;
    /**
     * The four edges this cell draws, already resolved against the table.
     *
     * Resolved here rather than by a renderer, because the answer depends on
     * where the cell sits in the grid — and a renderer holding a placed cell no
     * longer knows that.
     */
    readonly borders?: BoxBorders;
    readonly shadingFill?: string;
    /**
     * How tall the cell's own box is.
     *
     * Its row's height, unless a `w:vMerge` made it span several — which is why
     * a renderer must take it from HERE and not from the row it was drawn with.
     */
    readonly heightPx: number;
    /**
     * The cell's text. `paragraphIndex` and `lineIndex` are relative to the
     * CELL, not to the document — a line in a cell has no top-level paragraph.
     */
    readonly lines: readonly PlacedLine[];
    /** Tables nested inside the cell, already placed. */
    readonly rows: readonly PlacedRow[];
    /**
     * `w:textDirection`, where the cell's text was actually turned.
     *
     * Its lines are placed already transposed, so a renderer advances each
     * one along the page's Y and turns every operation it makes of it.
     */
    readonly turn?: 'ccw' | 'cw';
    /**
     * Boxes round the bordered paragraphs the cell holds — `w:pBdr` INSIDE a
     * cell, which draws just as it does on the page.
     *
     * Measured against LibreOffice: a bordered paragraph in a cell drew its
     * box across the cell's TEXT column, half a rule outside it either side,
     * and made the row two points taller for a one-point rule.
     */
    readonly paragraphBorders: readonly PlacedParagraphBorder[];
}

export interface PlacedRow {
    readonly yPx: number;
    readonly heightPx: number;
    readonly cells: readonly PlacedCell[];
    /** Index of the table among the document's blocks. */
    readonly blockIndex: number;
    readonly rowIndex: number;
    /** True when this is a header drawn again at the top of a later page. */
    readonly repeated: boolean;
}

/** A row measured at a set of column widths, before it is placed on a page. */
export interface MeasuredRow {
    readonly heightPx: number;
    readonly cells: readonly {
        xPx: number;
        widthPx: number;
        lines: PlacedLine[];
        /** Rows of any table nested in the cell, at the cell's own origin. */
        rows: PlacedRow[];
        /** What the cell's own content needs. */
        heightPx: number;
        /**
         * The box the cell DRAWS in: its row's height, or the whole span's
         * once a merge has been resolved. Settled by
         * {@link resolveTableGeometry}, which is the only thing that can see
         * every row at once.
         */
        boxHeightPx: number;
        /** How many rows the cell covers; 0 for one a merge has swallowed. */
        rowSpan: number;
        /** `w:tcMar` resolved against the table, so nothing downstream re-asks. */
        margins: ResolvedCellMargins;
        /**
         * Set where this cell's text is TURNED, and only where the turn could
         * be laid out — a cell with no upright neighbour to set the row's
         * height keeps its text upright and carries nothing here.
         */
        turn?: 'ccw' | 'cw';
        /** Boxes round the bordered paragraphs the cell holds. */
        paragraphBorders: PlacedParagraphBorder[];
        borders?: BoxBorders;
        shadingFill?: string;
    }[];
}

/** A table sized and placed against the column it landed in. */
export interface FittedTable {
    /** The same table, its grid scaled to the width it asked for. */
    readonly table: Table;
    /** How far right of the column's own left edge the table starts. */
    readonly offsetPx: number;
}

/**
 * Size a table against the column it lands in, and say where it starts.
 *
 * `w:tblW` scales the WHOLE grid proportionally rather than trimming the last
 * column: LibreOffice printed a two-column 4000+4000 grid at 200pt apiece when
 * asked for half a 538.58pt page, and at 100 apiece when asked for 4000 twips.
 */
export function fitTable(table: Table, availableWidthPx: number): FittedTable {
    const grid = table.columnWidthsPx.reduce((sum, width) => sum + width, 0);
    const asked = undefined === table.preferredWidthFraction
        ? table.preferredWidthPx
        : availableWidthPx * table.preferredWidthFraction;

    const widthPx = undefined === asked || asked <= 0 || grid <= 0 ? grid : asked;
    const scale = widthPx / grid;
    const fitted = grid > 0 && 1 !== scale
        ? { ...table, columnWidthsPx: table.columnWidthsPx.map((width) => width * scale) }
        : table;

    // An indent counts only where the table is left-aligned: measured against
    // LibreOffice, a centred table with half an inch of indent printed in
    // exactly the same place as one with none.
    if ('center' === table.alignment) {
        return { table: fitted, offsetPx: (availableWidthPx - widthPx) / 2 };
    }
    if ('right' === table.alignment) {
        return { table: fitted, offsetPx: availableWidthPx - widthPx };
    }

    // A table that simply starts at the margin hangs its left border half
    // outside it — measured at two widths, a 2pt border shifted the table 1pt
    // and a 6pt border 3. A table given a `w:tblInd` does NOT: that indent is
    // measured to the CELL'S TEXT, so the border lands a cell margin left of
    // the number and the hang is gone. Half an inch of indent against a 5.75pt
    // cell margin put the border at 30.25pt and the text at the whole 36.
    //
    // And it is the LEADING CELL'S margin, not the table's: with the first
    // cell overriding its own to 25pt, the same indent put the border at 39.4
    // rather than 58.6 — and the text still landed on margin plus indent.
    const leadingPx = table.rows[0]?.cells[0]?.margins?.leftPx ?? table.cellMarginLeftPx;

    return {
        table: fitted,
        // Subtracted from nought rather than negated, so a table with no
        // border of its own comes back at +0 and not the -0 that negating
        // zero gives.
        offsetPx: undefined === table.indentPx
            ? 0 - ruleLeft(table) / 2
            : table.indentPx - leadingPx,
    };
}

/** A cell's four margins, its own where it has them and the table's elsewhere. */
export function resolveCellMargins(table: Table, cell: TableCell): ResolvedCellMargins {
    return {
        leftPx: cell.margins?.leftPx ?? table.cellMarginLeftPx,
        rightPx: cell.margins?.rightPx ?? table.cellMarginRightPx,
        topPx: cell.margins?.topPx ?? table.cellMarginTopPx,
        bottomPx: cell.margins?.bottomPx ?? table.cellMarginBottomPx,
    };
}

/**
 * The width of the rule standing above row `index`, which is room the flow
 * has to keep clear.
 *
 * Measured against LibreOffice: the rule is CENTRED in a gap of its own width,
 * so the content above it and the content below are a whole width apart.
 *
 * ## The room is the rule that is DRAWN — so a cell's own beats the table's
 *
 * This asked the table first and took the widest of it and the cells, and a
 * cell cannot declare `insideH` at all, so the table's number always won. A
 * document whose cells state a hairline inside a table declaring half a point
 * therefore drew an eighth of a point and reserved half of one — every row, all
 * the way down. Found by printing a real document and diffing it
 * against ourselves: its rows stepped 10.85 where LibreOffice stepped 10.45.
 *
 * Each cell answers for ITSELF, falling back to the table only where it is
 * silent, and the widest claim on the edge wins. Measured, four tables of three
 * rows, room read off consecutive baselines less the 10.35 line:
 *
 *   * table 1/2pt, cells 1/8pt   -> 0.10   the cell's, thinner than the table's
 *   * table 1/2pt, cells silent  -> 0.50   the table's, nothing overriding it
 *   * table 1/8pt, cells 3pt     -> 3.00   the cell's, thicker than the table's
 *   * cells 1/8pt and 3pt        -> 3.00   the widest of the two
 *
 * Both SIDES of the edge count. A row stating 3pt under itself above a row
 * stating an eighth over itself took 3.00 — the fifth table, which is the only
 * pair the other four cannot separate, since their cells state four equal
 * sides. That is the same widest-wins the renderer resolves the shared edge
 * with, which is what keeps the room and the rule in step.
 */
export function ruleAbove(table: Table, index: number): number {
    const fallbackPx = (0 === index ? table.borders?.top : table.borders?.insideH)?.widthPx ?? 0;
    const claims = (table.rows[index]?.cells ?? [])
        .map((cell) => cell.borders?.top?.widthPx ?? fallbackPx);

    if (0 !== index) {
        claims.push(...(table.rows[index - 1]?.cells ?? [])
            .map((cell) => cell.borders?.bottom?.widthPx ?? fallbackPx));
    }

    return Math.max(0, ...claims);
}

/**
 * The rule down the table's LEFT edge, which is how far it hangs into the
 * margin — see {@link fitTable}.
 *
 * A cell's own border beats the table's, as it does for {@link ruleAbove}: a
 * table whose CELLS declare the border and which itself declares none still
 * hangs. Measured with the border stated only on the cells, at two
 * widths — a 1pt rule printed centred on 71.50 and a 3pt one on 70.50, each
 * putting its outer edge on the 72pt margin.
 *
 * But NOT the widest, which is where `ruleAbove`'s rule stops applying and
 * where guessing by analogy went wrong. A table whose first row claims 1pt and
 * whose second claims 3pt printed BOTH rules centred on 71.50: the hang is the
 * first row's, and a thicker rule below simply spills either side of the same
 * edge. A horizontal rule is shared between the two rows it divides and has to
 * be negotiated; this one is the table's own frame, and the first row states it.
 */
export function ruleLeft(table: Table): number {
    return table.rows[0]?.cells[0]?.borders?.left?.widthPx
        ?? table.borders?.left?.widthPx
        ?? 0;
}

/**
 * The same for the rule below the last row.
 *
 * Measured in the same print: a table declaring half a point whose cells state
 * an eighth drew its BOTTOM edge at the eighth as well, so the outer rules
 * follow the cell exactly as the inside ones do.
 */
export function ruleBelow(table: Table): number {
    const fallbackPx = table.borders?.bottom?.widthPx ?? 0;
    const last = table.rows[table.rows.length - 1];

    return Math.max(0, ...(last?.cells ?? [])
        .map((cell) => cell.borders?.bottom?.widthPx ?? fallbackPx));
}

/** Which row this is within its table — what picks outer borders from inside ones. */
export interface RowPosition {
    readonly firstRow: boolean;
    readonly lastRow: boolean;
}

/**
 * Where each grid column starts, and how wide a cell spanning from it is.
 *
 * Cumulative rather than per-cell because `w:gridSpan` lets one cell cover
 * several columns: the widths come from the table's grid and the cells only say
 * how many of them they take.
 */
export function columnOffsets(widths: readonly number[]): number[] {
    const offsets: number[] = [];
    let x = 0;
    for (const width of widths) {
        offsets.push(x);
        x += width;
    }

    return offsets;
}

/**
 * Measure one row: lay every cell out in its own column and take the tallest.
 *
 * The tallest rather than the first, because a row is as tall as its fullest
 * cell — sizing it from one cell puts the row's own bottom border through the
 * text of another.
 */
/**
 * How many rows each cell of a table covers.
 *
 * One for an ordinary cell, N for the top of a merge, and ZERO for a cell a
 * merge from above has swallowed. Zero rather than absent because such a cell
 * still holds its place in the row: dropping it would shift every cell after it
 * one grid column to the left.
 *
 * A merge is matched by the grid COLUMN it starts in, not by its position
 * among the row's cells — a `w:gridSpan` earlier in one row and not the next
 * would otherwise pair up cells that are nowhere near each other.
 */
export function verticalSpans(table: Table): number[][] {
    const spans = table.rows.map((row) => row.cells.map(() => 1));
    // Column -> where the merge that is open there began.
    const open = new Map<number, { row: number; cell: number }>();

    table.rows.forEach((row, rowIndex) => {
        const seen = new Set<number>();
        let column = 0;

        row.cells.forEach((cell, cellIndex) => {
            const at = column;
            column += Math.max(1, cell.gridSpan);
            seen.add(at);

            const above = open.get(at);
            if ('continue' === cell.verticalMerge && undefined !== above) {
                spans[above.row]![above.cell] = rowIndex - above.row + 1;
                spans[rowIndex]![cellIndex] = 0;

                return;
            }

            open.delete(at);
            if ('restart' === cell.verticalMerge) {
                open.set(at, { row: rowIndex, cell: cellIndex });
            }
        });

        // A column this row said nothing about cannot go on being merged into.
        for (const at of [...open.keys()]) {
            if (!seen.has(at)) {
                open.delete(at);
            }
        }
    });

    return spans;
}

/** How much of a row's spare height sits ABOVE the cell's content. */
const ALIGNMENT_SHARE: Record<CellVerticalAlign, number> = {
    top: 0,
    center: 0.5,
    bottom: 1,
};

/**
 * The same blocks, with every inline picture turned into its line's frame.
 *
 * Measured over five sizes: ALONG a turned line a picture reserves
 * its HEIGHT — 18x9 and 36x9 both reserved 9.00, 18x18 and 9x18 both 18.00, so
 * doubling the width moved nothing. In the line's own frame its sides are the
 * other way round, and exchanging them here is what lets the line breaker
 * measure a turned line exactly as it measures an upright one.
 *
 * ACROSS the line it does not stand on the baseline, which is the other half
 * and the one the exchange alone cannot say. A picture that fits is
 * CENTRED on the line's ascent; one too long for its line takes a line of its
 * own and the text hangs below it. Both are placements of the picture within
 * its line BOX, so the picture is only labelled here — the arithmetic wants the
 * paragraph's own ascent and belongs where the box is built.
 */
function turnPictures(blocks: readonly Block[], availablePx: number): readonly Block[] {
    return blocks.map((block) => {
        if (isTable(block)
            || !block.runs.some((run) => undefined !== run.image || undefined !== run.shape)) {
            return block;
        }

        return {
            ...block,
            runs: block.runs.map((run) => {
                // A drawn shape turns with the line the same way a picture
                // does, at least along it: measured, a 36x18 shape in
                // a turned cell started the text after it 18.05 above the row's
                // foot — its HEIGHT, where we charged the whole 36.00 of its
                // width. ACROSS the line it printed 31.70 from the cell's edge,
                // which is neither the 36.00 of standing on the baseline nor
                // the 22.69 of a picture's centring, and one measurement is not
                // a rule; standing is the closer of the two and is left alone.
                if (undefined !== run.shape) {
                    // The wrap distance a shape keeps is per-SIDE, and this
                    // line now advances DOWN the page, where the top and
                    // bottom distances apply rather than the left and right.
                    // Those default to nought, which is exactly what was
                    // measured: 18.05 charged for a 36x18 shape, its height
                    // and not a point more. Carrying the across-the-line
                    // figure through the turn charged 18 too much and the
                    // suite caught it the day it was built.
                    const alongPx = run.shape.heightPx > availablePx ? 0 : run.shape.heightPx;

                    return {
                        ...run,
                        shape: {
                            ...run.shape,
                            widthPx: alongPx,
                            heightPx: run.shape.widthPx,
                            advanceWidthPx: alongPx,
                        },
                    };
                }

                if (undefined === run.image) {
                    return run;
                }

                // A picture too long for the line is not charged for at all,
                // and it is not on the text's line either. Measured: 36x36 in
                // a 35.5pt row started its text 0.60 above the foot, exactly
                // like a line with no picture, and 45.40 across from the
                // cell's edge — the whole picture, then the text's own ascent
                // under it. LibreOffice gives a picture that cannot fit a line
                // of its own rather than reserving what it has, and reserving
                // the whole 36.00 along puts the text outside its own row.
                //
                // `>` and `>=` are indistinguishable to every fixture and a
                // mutation between them survives: the boundary is a picture
                // exactly as tall as the line, and the line's length is the
                // row's height, which is itself derived from the upright
                // cells' text. No fixture can be authored to sit ON that, only
                // just either side, so the probe would answer about the
                // fixture rather than the rule.
                const fits = run.image.heightPx <= availablePx;

                return {
                    ...run,
                    image: {
                        ...run.image,
                        widthPx: fits ? run.image.heightPx : 0,
                        heightPx: run.image.widthPx,
                        acrossLine: fits ? 'centred' as const : 'stacked' as const,
                    },
                };
            }),
        };
    });
}

export function measureRow(
    row: TableRow,
    table: Table,
    position: RowPosition,
    measure: (blocks: readonly Block[], widthPx: number) => StackedBlocks,
    spans?: readonly number[],
): MeasuredRow {
    const offsets = columnOffsets(table.columnWidthsPx);
    const cells: MeasuredRow['cells'][number][] = [];

    /**
     * How tall this row is before any TURNED cell is measured.
     *
     * A turned cell's lines run along the row's HEIGHT, so it cannot be
     * measured until the height is known — and the height comes from the cells
     * that are NOT turned. Measured against LibreOffice: a turned
     * cell beside a neighbour of three paragraphs put all six of its glyphs on
     * ONE line 35.5pt long, which is the neighbour's height exactly. So the
     * order is sequential, not circular.
     *
     * A turned cell with no upright neighbour has nothing to take a height
     * from. LibreOffice declines to lay such a table out at all, so
     * there is no answer to copy and this leaves the cell upright — which is
     * what it did before `w:textDirection` was read.
     */
    const uprightHeightPx = Math.max(
        row.heightPx ?? 0,
        ...row.cells
            .filter((cell) => undefined === cell.textDirection)
            .map((cell, cellIndex) => {
                const span = Math.max(1, cell.gridSpan);
                const start = row.cells.slice(0, cellIndex)
                    .reduce((total, before) => total + Math.max(1, before.gridSpan), 0);
                let widthPx = 0;
                for (let index = start; index < start + span; index++) {
                    widthPx += table.columnWidthsPx[index] ?? 0;
                }
                const margins = resolveCellMargins(table, cell);

                return measure(cell.paragraphs, Math.max(1, widthPx - margins.leftPx
                    - margins.rightPx)).heightPx + margins.topPx + margins.bottomPx;
            }),
    );

    let column = 0;
    row.cells.forEach((cell, cellIndex) => {
        const rowSpan = spans?.[cellIndex] ?? 1;
        const span = Math.max(1, cell.gridSpan);
        let widthPx = 0;
        for (let index = column; index < column + span; index++) {
            widthPx += table.columnWidthsPx[index] ?? 0;
        }

        const margins = resolveCellMargins(table, cell);
        // A TURNED cell's lines run along the row's height, so that — and not
        // the column's width — is the length they break at. With no upright
        // neighbour the height is nought and the cell stays upright.
        const turned = undefined !== cell.textDirection && uprightHeightPx > 0;
        const inner = turned
            ? uprightHeightPx - margins.topPx - margins.bottomPx
            : widthPx - margins.leftPx - margins.rightPx;
        // Measured against LibreOffice, text in a continued cell was not drawn
        // at all — so there is no reason to break it into lines. Skipping the
        // work is all this does: the cell is left out of the placed row, and
        // its box is nil, either of which would drop the lines anyway.
        const laid: StackedBlocks = 0 === rowSpan
            ? { lines: [], rows: [], borders: [], heightPx: 0 }
            : measure(
                turned ? turnPictures(cell.paragraphs, inner) : cell.paragraphs,
                Math.max(1, inner),
            );

        const borders = resolveCellBorders(table.borders, cell.borders, {
            ...position,
            firstColumn: 0 === column,
            // The LAST grid column this cell covers, not the first: a cell
            // spanning to the edge of the table draws the table's right border.
            lastColumn: column + span >= table.columnWidthsPx.length,
        });
        const fill = cell.shadingFill ?? table.shadingFill;

        // A turned cell does not GROW the row: measured, the row stayed the
        // height its upright neighbour asked for. What its own text needs is
        // width, and it takes that from the column it was given.
        const heightPx = 0 === rowSpan
            ? 0
            : turned
                ? uprightHeightPx
                : laid.heightPx + margins.topPx + margins.bottomPx;

        cells.push({
            margins,
            ...(turned ? { turn: 'btLr' === cell.textDirection ? 'ccw' : 'cw' } : {}),
            xPx: offsets[column] ?? 0,
            widthPx,
            lines: laid.lines,
            rows: laid.rows,
            paragraphBorders: laid.borders,
            heightPx,
            boxHeightPx: heightPx,
            rowSpan,
            borders,
            ...(undefined === fill ? {} : { shadingFill: fill }),
        });

        column += span;
    });

    // An empty row still has height, or the borders of the rows either side of
    // it would coincide. A cell that does not stand in exactly one row is left
    // out: a merged one is measured against its whole SPAN, which this cannot
    // see, and a swallowed one holds nothing.
    const contentPx = cells.reduce(
        (tallest, cell) => 1 === cell.rowSpan ? Math.max(tallest, cell.heightPx) : tallest,
        0,
    );

    // Neighbours share the edge between them, so it is settled once and both
    // are given the answer.
    for (let index = 1; index < cells.length; index++) {
        const left = cells[index - 1]!;
        const right = cells[index]!;
        const winner = strongerBorder(left.borders?.right, right.borders?.left);

        left.borders = withSide(left.borders, 'right', winner);
        right.borders = withSide(right.borders, 'left', winner);
    }

    return { heightPx: rowHeight(row, contentPx), cells };
}

/** A row's own height, once its cells have been measured. */
function rowHeight(row: TableRow, contentPx: number): number {
    const asked = row.heightPx;

    return undefined === asked
        ? contentPx
        : ('exact' === row.heightRule ? asked : Math.max(contentPx, asked));
}

/**
 * Settle every row's height and every cell's box, with the whole table in hand.
 *
 * Vertical merging is why this cannot happen a row at a time. A merged cell is
 * measured against the SPAN it covers, so its own row does not know how tall it
 * has to be — and neither does any row until every merge has had its say.
 *
 * Measured against LibreOffice: the rows a merge covers keep their own natural
 * heights, and where the merged content is taller than all of them together the
 * shortfall lands on the LAST row of the span. Five lines merged across two
 * one-line rows left the first at one line and grew the second to four.
 */
export function resolveTableGeometry(
    table: Table,
    measured: readonly MeasuredRow[],
    spans: readonly (readonly number[])[],
): number[] {
    const heights = measured.map((row) => row.heightPx);

    // A merge that ends higher up is settled first: growing a row it covers
    // changes what a merge reaching PAST that row still needs.
    const merges: { row: number; cell: number; span: number }[] = [];
    spans.forEach((rowSpans, rowIndex) => {
        rowSpans.forEach((span, cellIndex) => {
            if (span > 1) {
                merges.push({ row: rowIndex, cell: cellIndex, span });
            }
        });
    });
    merges.sort((one, other) => (one.row + one.span) - (other.row + other.span));

    /**
     * How tall a merge reaching over `count` rows from `from` stands.
     *
     * The rows it covers AND the rules between them: measured against
     * LibreOffice, a cell merged over three 11.5pt rows drew a box 37.5pt tall
     * — 34.5 of text, the two one-point inside rules its span crosses, and
     * half a rule outside at each end. Summing the rows alone leaves it short
     * by one rule per row it swallows, and the box stops above its own foot.
     */
    const spanned = (from: number, count: number): number =>
        heights.slice(from, from + count).reduce((total, height) => total + height, 0)
        + Array.from({ length: Math.max(0, count - 1) },
            (_, index) => ruleAbove(table, from + index + 1))
            .reduce((total, width) => total + width, 0);

    for (const merge of merges) {
        const needs = measured[merge.row]?.cells[merge.cell]?.heightPx ?? 0;
        const last = merge.row + merge.span - 1;
        const shortfall = needs - spanned(merge.row, merge.span);
        if (shortfall > 0) {
            heights[last] = (heights[last] ?? 0) + shortfall;
        }
    }

    measured.forEach((row, rowIndex) => {
        row.cells.forEach((cell) => {
            cell.boxHeightPx = 0 === cell.rowSpan ? 0 : spanned(rowIndex, cell.rowSpan);
        });
    });

    alignCells(table, measured);

    return heights;
}

/**
 * Move each cell's content down inside the box it was given, and drop whatever
 * still hangs out of the bottom.
 */
function alignCells(table: Table, measured: readonly MeasuredRow[]): void {
    measured.forEach((row, rowIndex) => {
        const declared = table.rows[rowIndex]?.cells ?? [];

        row.cells.forEach((cell, cellIndex) => {
            // Negative slack is an `exact` row smaller than its own text, where
            // there is nothing to distribute: measured against LibreOffice,
            // bottom alignment left such a cell exactly where top did.
            const slack = Math.max(0, cell.boxHeightPx - cell.heightPx);
            const shift = slack * ALIGNMENT_SHARE[declared[cellIndex]?.verticalAlign ?? 'top'];

            // A line is drawn when its TOP edge is still inside the box, which
            // is how LibreOffice decided: an exact row 15px tall drew the
            // second line of three, whose top was inside it, and dropped the
            // third. Only an `exact` row can lose a line this way — every other
            // rule leaves the box at least as tall as the cell, so the filter
            // passes everything.
            //
            // A TURNED cell stacks its lines ACROSS the cell, so the room they
            // run out of is the WIDTH. Measured: LibreOffice ran a
            // turned cell to seventeen lines over a 200pt column and off its
            // right edge, in a row 13.2pt tall — asking the row's height about
            // the cross axis dropped all but the first two, and text vanished
            // from cells that had the room for it.
            const roomPx = undefined === cell.turn
                ? cell.boxHeightPx - cell.margins.topPx
                : cell.widthPx - cell.margins.leftPx - cell.margins.rightPx;

            cell.lines = cell.lines
                .filter((line) => line.yPx + shift < roomPx)
                .map((line) => 0 === shift ? line : { ...line, yPx: line.yPx + shift });
        });
    });
}

/**
 * The box ONE line got.
 *
 * Per line rather than per paragraph because a picture makes its own line
 * taller and no other — measured against LibreOffice.
 */
export interface LineMetrics {
    readonly heightPx: number;
    /** How far below the box's top the baseline sits. */
    readonly baselinePx: number;
}

/** How a caller breaks one paragraph into lines at a given width. */
export type MeasureParagraph = (paragraph: Paragraph, availableWidth: number)
    => {
        lines: Line[];
        /**
         * What a line of this paragraph is worth before the breaking has said
         * which runs landed where. Only a float uses it, and it has to: the
         * lines a float displaces are decided before there are any.
         */
        lineHeight: number;
        /** One per line, in the same order. */
        metrics: readonly LineMetrics[];
    };

/** A column of blocks, stacked with no pagination. */
export interface StackedBlocks {
    readonly lines: PlacedLine[];
    /** Boxes round the bordered paragraphs among them, at the column's origin. */
    readonly borders: PlacedParagraphBorder[];
    /** Rows of every table among the blocks, at the column's own origin. */
    readonly rows: PlacedRow[];
    readonly heightPx: number;
}

/**
 * Stack blocks into a column of a given width, with no pagination.
 *
 * This is what a table cell is, and what a header is: a little page of
 * unbounded height. Sharing the paragraph measurement with the page flow is
 * what keeps a cell's line breaking identical to the same text outside a table.
 *
 * A table among the blocks is laid out here too, which is what makes nesting
 * work — a cell stacks its blocks with this, and one of those blocks may be a
 * table whose cells stack theirs with it in turn.
 */
export function stackBlocks(
    blocks: readonly Block[],
    widthPx: number,
    measureParagraph: MeasureParagraph,
): StackedBlocks {
    const lines: PlacedLine[] = [];
    const rows: PlacedRow[] = [];
    const borders: PlacedParagraphBorder[] = [];
    let y = 0;

    /**
     * The box being built, if a bordered paragraph is open.
     *
     * The same run-of-paragraphs rule the page flow follows, and simpler here:
     * a cell has no page breaks, so a box opened inside one always closes
     * inside it too and is never split.
     */
    let open: { box: PlacedParagraphBorder; style: ParagraphStyle } | null = null;

    /**
     * What the block above already spent below itself — see the page flow,
     * where the same rule is measured. A cell stacks paragraphs the
     * same way the page does, and two of them collapse their spacing here too.
     */
    let spentBelowPx = 0;

    const closeBox = (): void => {
        if (null === open) {
            return;
        }

        borders.push(open.box);
        y += borderRoom(open.box.borders.bottom);
        open = null;
    };

    blocks.forEach((block, blockIndex) => {
        if (isTable(block)) {
            y += block.spaceBeforePx;
            spentBelowPx = 0;
            const { table: fitted, offsetPx } = fitTable(block, widthPx);
            const { measured, heights } = measureTable(fitted, measureParagraph);

            measured.forEach((row, rowIndex) => {
                const heightPx = heights[rowIndex] ?? row.heightPx;
                rows.push(placeRow(row, fitted, { xPx: offsetPx, yPx: y }, heightPx, {
                    blockIndex,
                    rowIndex,
                    repeated: false,
                }));
                y += heightPx;
            });

            y += block.spaceAfterPx;

            return;
        }

        if (isStructuralTail(blocks, blockIndex)) {
            return;
        }

        const paragraph = block;
        const paragraphIndex = blockIndex;
        const style: ParagraphStyle = paragraph.style;

        // A paragraph with no border of its own ends the box the one before it
        // opened; one carrying the same border extends it. Measured against
        // LibreOffice, a bordered paragraph in a cell drew its box across the
        // cell's TEXT column and made the row two points taller for a
        // one-point rule — the same room a paragraph takes on the page.
        if (undefined === style.borders
            || null === open
            || !sameBorders(open.box.borders, style.borders)) {
            closeBox();
        } else if (undefined !== style.borders.insideH) {
            // The box goes on, with a `w:between` rule across it — half the
            // rule below the text above, with the whole of its width out of
            // the flow. Measured in a cell as on the page: the pair stepped
            // 14.5pt for a three-point rule where an unruled pair steps 11.5.
            open = {
                ...open,
                box: {
                    ...open.box,
                    innerYPx: [
                        ...open.box.innerYPx,
                        y + style.borders.insideH.widthPx / 2,
                    ],
                },
            };
            y += borderRoom(style.borders.insideH);
        }

        // Opening a box costs its top rule's room before the first line.
        if (undefined !== style.borders && null === open) {
            y += borderRoom(style.borders.top);
        }
        const available = widthPx - style.indentLeftPx - style.indentRightPx;
        const measured = measureParagraph(paragraph, available);


        y += Math.max(0, style.spaceBeforePx - spentBelowPx);

        measured.lines.forEach((line, lineIndex) => {
            const indent = 0 === lineIndex ? style.indentFirstLinePx : 0;
            const { offsetPx, wordSpacingPx } = alignLine(
                line,
                available - indent,
                style.alignment,
                lineIndex === measured.lines.length - 1,
            );

            // Each line's own box: a picture makes the line it sits on taller
            // and leaves the rest of the paragraph alone.
            const box = measured.metrics[lineIndex]!;

            lines.push({
                ...(0 === lineIndex && undefined !== paragraph.marker
                    ? {
                        marker: {
                            run: paragraph.marker.run,
                            xPx: style.indentLeftPx - paragraph.marker.offsetPx,
                        },
                    }
                    : {}),
                line,
                xPx: style.indentLeftPx + indent + offsetPx,
                yPx: y,
                heightPx: box.heightPx,
                baselinePx: box.baselinePx,
                wordSpacingPx,
                paragraphIndex,
                lineIndex,
            });
            y += box.heightPx;

            if (undefined !== style.borders) {
                const side = style.borders;
                const edges = {
                    leftPx: style.indentLeftPx - borderStandoff(side.left),
                    rightPx: widthPx - style.indentRightPx + borderStandoff(side.right),
                    bottomPx: y + borderStandoff(side.bottom),
                };

                open = null === open
                    ? {
                        style,
                        box: {
                            ...edges,
                            borders: side,
                            topPx: y - box.heightPx - borderStandoff(side.top),
                            opensHere: true,
                            closesHere: true,
                            innerYPx: [],
                        },
                    }
                    : { style, box: { ...open.box, bottomPx: edges.bottomPx } };
            }
        });

        y += style.spaceAfterPx;
        spentBelowPx = style.spaceAfterPx;
    });

    closeBox();

    return { lines, rows, borders, heightPx: y };
}

/**
 * The empty paragraph OOXML makes a cell END with after a nested table.
 *
 * It is structure, not content: measured against LibreOffice it was drawn a
 * twentieth of a point tall, where an empty paragraph anywhere else in a cell
 * took a whole line. Giving it one would make every cell holding a table a
 * line taller than the document says.
 */
function isStructuralTail(blocks: readonly Block[], index: number): boolean {
    const block = blocks[index];

    return index === blocks.length - 1
        && index > 0
        && undefined !== block
        && !isTable(block)
        && block.runs.every((run) => '' === run.text)
        && isTable(blocks[index - 1]!);
}

/**
 * Every row of a table measured, with the heights its merges settled on.
 *
 * Shared by the page flow and by a cell, which need the same answer and would
 * otherwise each work it out — the page flow going on to break the rows across
 * pages, and a cell simply stacking them.
 */
export function measureTable(
    table: Table,
    measureParagraph: MeasureParagraph,
): { measured: MeasuredRow[]; spans: number[][]; heights: number[] } {
    const spans = verticalSpans(table);
    const measured = table.rows.map((row, rowIndex) => measureRow(
        row,
        table,
        { firstRow: 0 === rowIndex, lastRow: rowIndex === table.rows.length - 1 },
        (blocks, innerWidthPx) => stackBlocks(blocks, innerWidthPx, measureParagraph),
        spans[rowIndex],
    ));
    const heights = resolveTableGeometry(table, measured, spans);

    // The edge between two rows belongs to both of them, and a merged cell
    // meets the row below its SPAN rather than the one below its own — so the
    // cells are gathered by where their boxes END.
    for (let index = 1; index < measured.length; index++) {
        shareRowEdge(
            measured.flatMap((row, rowIndex) => row.cells.filter(
                (cell) => cell.rowSpan > 0 && rowIndex + cell.rowSpan === index,
            )),
            measured[index]!.cells.filter((cell) => cell.rowSpan > 0),
        );
    }

    return { measured, spans, heights };
}

/**
 * Settle the edge two stacked rows share, giving BOTH of them the winner.
 *
 * Both rather than clearing the loser, because a table that breaks across a
 * page has its upper row on one sheet and its lower row on the next: clearing
 * one would take the rule off whichever half did not keep it, and a continued
 * table would open with no line above it.
 *
 * Cells are paired by their x and width. A row whose cells do not line up with
 * the row above — different `w:gridSpan` on one of them — has no shared edge
 * this can identify, and both rows keep what they were given rather than having
 * a neighbour guessed for them.
 */
function shareRowEdge(
    upper: readonly MeasuredRow['cells'][number][],
    lower: readonly MeasuredRow['cells'][number][],
): void {
    for (const below of lower) {
        const above = upper.find(
            (cell) => cell.xPx === below.xPx && cell.widthPx === below.widthPx,
        );
        if (undefined === above) {
            continue;
        }

        const winner = strongerBorder(above.borders?.bottom, below.borders?.top);
        above.borders = withSide(above.borders, 'bottom', winner);
        below.borders = withSide(below.borders, 'top', winner);
    }
}

/**
 * A measured row cut in two at `atPx`, measured from the row's own top.
 *
 * What fits stays; what does not becomes a row of its own for the next page.
 * The cut falls between LINES — measured against LibreOffice, a six-line row
 * with four lines of room kept four and carried two over, rather than clipping
 * the fifth or moving the whole of it.
 *
 * The halves meet at a page boundary, so neither draws the edge between them:
 * the upper keeps its top and the lower its bottom, and a rule across the break
 * would be a line the document never asked for.
 */
export function splitRow(
    measured: MeasuredRow,
    atPx: number,
): { upper: MeasuredRow; lower: MeasuredRow } | null {
    // A cell a merge swallowed holds nothing to divide, and one that spans past
    // this row is not this row's to cut.
    if (measured.cells.some((cell) => 1 !== cell.rowSpan)) {
        return null;
    }

    // Each cell against its OWN top margin: a cell that pads itself has less
    // room in the same row than the one beside it.
    const kept = measured.cells.map((cell) => cell.lines.filter(
        (line) => line.yPx + line.heightPx <= atPx - cell.margins.topPx).length);

    // Nothing to gain: every line stays, or none does.
    if (kept.every((count, index) => count === measured.cells[index]!.lines.length)) {
        return null;
    }
    if (kept.every((count) => 0 === count)) {
        return null;
    }

    const upperHeight = Math.max(...measured.cells.map((cell, index) => {
        const last = cell.lines[kept[index]! - 1];

        return undefined === last ? 0 : last.yPx + last.heightPx + cell.margins.topPx;
    }));

    const half = (
        take: (cell: MeasuredRow['cells'][number], index: number) => PlacedLine[],
        side: 'top' | 'bottom',
        heightPx: number,
    ): MeasuredRow => ({
        heightPx,
        cells: measured.cells.map((cell, index) => ({
            ...cell,
            lines: take(cell, index),
            heightPx,
            boxHeightPx: heightPx,
            borders: withSide(cell.borders, side, undefined),
        })),
    });

    return {
        upper: half((cell, index) => cell.lines.slice(0, kept[index]), 'bottom', upperHeight),
        lower: half(
            (cell, index) => cell.lines.slice(kept[index]).map((line) => ({
                ...line,
                yPx: line.yPx - upperHeight + cell.margins.topPx,
            })),
            'top',
            measured.heightPx - upperHeight,
        ),
    };
}

/** One measured row, at the place the caller is putting it. */
export function placeRow(
    measured: MeasuredRow,
    table: Table,
    origin: { xPx: number; yPx: number },
    heightPx: number,
    identity: { blockIndex: number; rowIndex: number; repeated: boolean },
): PlacedRow {
    return {
        yPx: origin.yPx,
        heightPx,
        ...identity,
        // A cell a merge swallowed draws nothing at all: no text, no shading
        // and no border through the middle of the cell above.
        cells: measured.cells.filter((cell) => cell.rowSpan > 0).map((cell) => {
            // A TURNED cell's lines were broken along the row's height, so
            // they are placed across it, and the two turns are MIRROR IMAGES.
            // Measured: `btLr` runs each line UP the page from the cell's foot
            // and steps rightwards from its left edge; `tbRl` runs
            // DOWN from the head and steps leftwards from the right edge, its
            // first baseline at x=112.50 in a cell whose inner right edge is
            // 122.05. Taking one from the other by luck was the bug.
            if (undefined !== cell.turn) {
                const anticlockwise = 'ccw' === cell.turn;
                const leftPx = origin.xPx + cell.xPx + cell.margins.leftPx;
                const rightPx = origin.xPx + cell.xPx + cell.widthPx - cell.margins.rightPx;
                const footPx = origin.yPx + cell.boxHeightPx - cell.margins.bottomPx;
                const headPx = origin.yPx + cell.margins.topPx;

                return {
                    xPx: origin.xPx + cell.xPx,
                    widthPx: cell.widthPx,
                    heightPx: cell.boxHeightPx,
                    turn: cell.turn,
                    ...(undefined === cell.borders ? {} : { borders: cell.borders }),
                    ...(undefined === cell.shadingFill ? {} : { shadingFill: cell.shadingFill }),
                    // A quarter turn leaves a rectangle axis-aligned, so a
                    // paragraph's box is the same box with its axes exchanged
                    // — the sides it keeps clear are the ones the upright case
                    // already worked out. Measured: a boxed line in a
                    // turned cell printed 20.45 across by 44.45 along, which is
                    // the line's own 11.50 and the row's 35.50 with `w:space`
                    // 4pt and half a point of rule at every edge. Handed back
                    // EMPTY until this row, so the box was placed and never
                    // drawn — this arc's commonest defect.
                    paragraphBorders: cell.paragraphBorders.map((box) => ({
                        ...box,
                        leftPx: anticlockwise
                            ? leftPx + box.topPx
                            : rightPx - box.bottomPx,
                        rightPx: anticlockwise
                            ? leftPx + box.bottomPx
                            : rightPx - box.topPx,
                        topPx: anticlockwise ? footPx - box.rightPx : headPx + box.leftPx,
                        bottomPx: anticlockwise ? footPx - box.leftPx : headPx + box.rightPx,
                        // `w:between` rules run ACROSS a turned box, where the
                        // renderer draws them along. No measurement covers one
                        // — the fixture that found the box holds a single
                        // paragraph — so they are dropped rather than drawn the
                        // wrong way round.
                        innerYPx: [],
                    })),
                    // A nested table in a turned cell IS drawn by LibreOffice,
                    // text and rules both: measured, its `N1` printed
                    // turned at x=87.65 inside a box 13.50 across by 36.50
                    // along. We drop it, because turning a row means turning
                    // every cell, line and rule it holds — a tree, where a
                    // paragraph's box is one rectangle. Pinned in a test.
                    rows: [],
                    lines: cell.lines.map((line) => ({
                        ...line,
                        // The line's cross-axis offset becomes an x — always
                        // the LEFT edge of its box, so the renderer can find
                        // either baseline from it — and its start along the
                        // row's height becomes a y from the foot or the head.
                        xPx: anticlockwise
                            ? leftPx + line.yPx
                            : rightPx - line.yPx - line.heightPx,
                        yPx: anticlockwise ? footPx - line.xPx : headPx + line.xPx,
                    })),
                };
            }

            const dxPx = origin.xPx + cell.xPx + cell.margins.leftPx;
            const dyPx = origin.yPx + cell.margins.topPx;

            return {
                xPx: origin.xPx + cell.xPx,
                widthPx: cell.widthPx,
                heightPx: cell.boxHeightPx,
                paragraphBorders: cell.paragraphBorders.map((box) => ({
                    ...box,
                    leftPx: box.leftPx + dxPx,
                    rightPx: box.rightPx + dxPx,
                    topPx: box.topPx + dyPx,
                    bottomPx: box.bottomPx + dyPx,
                    innerYPx: box.innerYPx.map((value) => value + dyPx),
                })),
                ...(undefined === cell.borders ? {} : { borders: cell.borders }),
                ...(undefined === cell.shadingFill ? {} : { shadingFill: cell.shadingFill }),
                // The cell laid its content out from its own origin; this is
                // where that origin turned out to be.
                lines: cell.lines.map((line) => translateLine(line, dxPx, dyPx)),
                rows: cell.rows.map((row) => translateRow(row, dxPx, dyPx)),
            };
        }),
    };
}

export function translateLine(line: PlacedLine, dxPx: number, dyPx: number): PlacedLine {
    return {
        ...line,
        ...(undefined === line.marker
            ? {}
            : { marker: { ...line.marker, xPx: line.marker.xPx + dxPx } }),
        xPx: line.xPx + dxPx,
        yPx: line.yPx + dyPx,
    };
}

/** A placed row moved bodily, along with everything nested inside it. */
export function translateRow(row: PlacedRow, dxPx: number, dyPx: number): PlacedRow {
    return {
        ...row,
        yPx: row.yPx + dyPx,
        cells: row.cells.map((cell) => ({
            ...cell,
            xPx: cell.xPx + dxPx,
            lines: cell.lines.map((line) => translateLine(line, dxPx, dyPx)),
            rows: cell.rows.map((nested) => translateRow(nested, dxPx, dyPx)),
            paragraphBorders: cell.paragraphBorders.map((box) => ({
                ...box,
                leftPx: box.leftPx + dxPx,
                rightPx: box.rightPx + dxPx,
                topPx: box.topPx + dyPx,
                bottomPx: box.bottomPx + dyPx,
                innerYPx: box.innerYPx.map((value) => value + dyPx),
            })),
        })),
    };
}

/**
 * A copy of these borders with one side replaced.
 *
 * A copy because the resolved borders may be the very object another cell was
 * given: the sides are shared structurally, and writing through one would
 * change an edge nobody asked about.
 */
export function withSide(
    borders: BoxBorders | undefined,
    side: 'top' | 'bottom' | 'left' | 'right',
    value: BorderSide | undefined,
): BoxBorders {
    const out: Record<string, BorderSide> = { ...borders };
    if (undefined === value) {
        delete out[side];
    } else {
        out[side] = value;
    }

    return out;
}

/** The leading rows marked as headers — the ones that repeat. */
export function headerRows(table: Table): TableRow[] {
    const out: TableRow[] = [];
    for (const row of table.rows) {
        if (!row.isHeader) {
            break;
        }
        out.push(row);
    }

    return out;
}
