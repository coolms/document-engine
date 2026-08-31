/**
 * Table borders and shading.
 *
 * A table's edges are not a property of the table: they are a property of every
 * CELL, arrived at by asking the cell first and the table second, and the answer
 * depends on where in the grid the cell sits. A border declared on the table as
 * `insideH` belongs to the cells that have a neighbour below, and the one
 * declared as `bottom` belongs only to the cells on the last row.
 *
 * Its own module for the reason `alignment` and `baseline` are: `table-layout`
 * needs the resolution at runtime and imports `page-layout` for TYPES ONLY, so
 * a shared function cannot live there.
 */

/**
 * How a border is drawn.
 *
 * OOXML has around thirty `w:val` styles, most of them decorative Word 6
 * survivals — wavy, three-dimensional, and several kinds of art border. They
 * are folded onto these four: the WIDTH and the presence of a border move the
 * text and the eye, and the difference between `thinThickSmallGap` and
 * `thickThinSmallGap` does not.
 */
export type BorderStyle = 'solid' | 'dashed' | 'dotted' | 'double';

export interface BorderSide {
    readonly widthPx: number;
    /**
     * `w:space` — how far the rule stands OUTSIDE what it surrounds.
     *
     * A paragraph's border is what needs it: measured against LibreOffice, six
     * points of space moved the box six points further out on every side and
     * left the text exactly where it was. A table states it too and does not
     * use it.
     */
    readonly spacePx?: number;
    /** `#RRGGBB`. `w:color="auto"` resolves to black, which is what Word draws. */
    readonly colorHex: string;
    readonly style: BorderStyle;
}

/**
 * The borders declared on one table or one cell.
 *
 * `insideH` and `insideV` are meaningful on a TABLE only — a cell has no
 * inside. They are kept in the same shape because both are read from the same
 * `w:*Borders` element, and a second type would only differ by two fields.
 */
export interface BoxBorders {
    readonly top?: BorderSide;
    readonly left?: BorderSide;
    readonly bottom?: BorderSide;
    readonly right?: BorderSide;
    /** Between a row and the one below it. */
    readonly insideH?: BorderSide;
    /** Between a column and the one beside it. */
    readonly insideV?: BorderSide;
}

/** Where a cell sits in the grid, which is what picks between outer and inside. */
export interface CellPosition {
    readonly firstRow: boolean;
    readonly lastRow: boolean;
    readonly firstColumn: boolean;
    readonly lastColumn: boolean;
}

/**
 * The four edges this cell actually draws.
 *
 * The cell's own declaration wins on every side. Where it declares nothing, an
 * edge on the outside of the grid takes the table's matching outer border and
 * an edge with a neighbour takes the table's inside one.
 *
 * ## Shared edges are drawn twice, but they now AGREE
 *
 * Two cells side by side both draw the boundary between them. The pair is
 * resolved first — see {@link strongerBorder} — and both are given the winner,
 * so the second lands exactly on the first and nothing shows.
 *
 * Both rather than one, because a table that breaks across a page has an upper
 * row on one sheet and a lower row on the next: clearing the loser would take
 * the rule off whichever half did not keep it, and a continued table would open
 * with no line above it.
 */
export function resolveCellBorders(
    table: BoxBorders | undefined,
    cell: BoxBorders | undefined,
    at: CellPosition,
): BoxBorders {
    const side = (
        own: BorderSide | undefined,
        outer: BorderSide | undefined,
        inside: BorderSide | undefined,
        isOutside: boolean,
    ): BorderSide | undefined => own ?? (isOutside ? outer : inside);

    return pick({
        top: side(cell?.top, table?.top, table?.insideH, at.firstRow),
        bottom: side(cell?.bottom, table?.bottom, table?.insideH, at.lastRow),
        left: side(cell?.left, table?.left, table?.insideV, at.firstColumn),
        right: side(cell?.right, table?.right, table?.insideV, at.lastColumn),
    });
}

/**
 * Which of two borders meeting on one edge is drawn.
 *
 * The HEAVIER wins. On a tie the LATER one does — the right-hand cell's left
 * border, the lower row's top border.
 *
 * Measured out of LibreOffice rather than reasoned about: a 4pt rule beats a
 * 1pt one whichever side declares it, and two 2pt rules of different colours
 * resolve to the right-hand cell's — the same pair with the colours SWAPPED
 * resolves the other way, which is what rules out a rule about colour.
 *
 * Word documents a longer chain — weight, then style, then colour — and only
 * the weight and the tie-break are implemented. Two borders of equal weight and
 * different STYLE therefore resolve by position where Word would rank the
 * style; stated rather than guessed at.
 */
export function strongerBorder(
    earlier: BorderSide | undefined,
    later: BorderSide | undefined,
): BorderSide | undefined {
    if (undefined === earlier) {
        return later;
    }
    if (undefined === later) {
        return earlier;
    }

    return earlier.widthPx > later.widthPx ? earlier : later;
}

/** Drop the sides that resolved to nothing, so an absent side stays absent. */
function pick(sides: Record<string, BorderSide | undefined>): BoxBorders {
    const out: Record<string, BorderSide> = {};
    for (const [name, value] of Object.entries(sides)) {
        if (undefined !== value) {
            out[name] = value;
        }
    }

    return out;
}

/**
 * How far outside its edge a side's rule is CENTRED, `w:space` included.
 *
 * Measured against LibreOffice: a one-point box round an 11.5pt line came out
 * 12.45pt between rule centres, and six points of `w:space` moved every edge
 * six points further out while the text stayed put.
 */
export function borderStandoff(side: BorderSide | undefined): number {
    return undefined === side ? 0 : (side.spacePx ?? 0) + side.widthPx / 2;
}

/**
 * The room a side takes outside the text: its space plus its whole width.
 *
 * The rule is centred half a width outside the box, so its OUTER edge — what
 * the text above or below has to clear — is a full width out. An 11.5pt line
 * stepped 12.5 under a one-point border and 18.5 with six points of space.
 */
export function borderRoom(side: BorderSide | undefined): number {
    return undefined === side ? 0 : (side.spacePx ?? 0) + side.widthPx;
}

/**
 * Whether two paragraphs carry the SAME border, and so share one box.
 *
 * By value rather than by identity: two paragraphs of one style hold the same
 * object, but two that each state the same border do not, and LibreOffice drew
 * one outline round both either way.
 */
export function sameBorders(one: BoxBorders, other: BoxBorders): boolean {
    const sides = ['top', 'left', 'bottom', 'right', 'insideH'] as const;

    return sides.every((side) => {
        const a = one[side];
        const b = other[side];

        return undefined === a || undefined === b
            ? a === b
            : a.widthPx === b.widthPx && a.colorHex === b.colorHex
                && a.style === b.style && (a.spacePx ?? 0) === (b.spacePx ?? 0);
    });
}

/**
 * A paragraph border, as the box it is drawn as.
 *
 * ONE box for a run of paragraphs, not one each: measured against LibreOffice,
 * two paragraphs carrying the same border came out as a single outline with no
 * rule between them. The edges are the line box's own, with the rule centred
 * half a width outside them and `w:space` pushing it further out again.
 */
export interface PlacedParagraphBorder {
    /** The box the rules are centred ON, `w:space` already taken in. */
    readonly leftPx: number;
    readonly rightPx: number;
    readonly topPx: number;
    readonly bottomPx: number;
    readonly borders: BoxBorders;
    /** False where the run began on an earlier page, and its top is up there. */
    readonly opensHere: boolean;
    /** False where it carries on overleaf. */
    readonly closesHere: boolean;
    /**
     * Where `w:between` puts a rule inside the box, one y per join.
     *
     * A merged run is ONE outline, so a rule between its paragraphs is inside
     * the box rather than a pair of edges — and LibreOffice drew it half a
     * width below the text above it, with the whole `w:space` beneath.
     */
    readonly innerYPx: readonly number[];
}
