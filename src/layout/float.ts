import type { InlineImage } from './image.js';
import type { Block, Paragraph, PlacedLine, PlacedRow } from './page-layout.js';

/**
 * Things that text flows around: a picture, or a text box.
 *
 * A float is not in the line. It is anchored to a PARAGRAPH and placed at
 * coordinates of its own, and the lines it overlaps step aside for it — so
 * unlike everything else in this engine, its position is decided before the
 * text near it is broken, and then decides how that text breaks.
 *
 * A text box is the same thing carrying WORDS instead of a picture. Everything
 * about the frame is shared, which is why the two are one shape here with one
 * field between them: the anchoring, the wrap and the clear space were built
 * for pictures and are right for boxes without a change.
 */

/** How text behaves where a float overlaps it. */
export type WrapMode =
    /** Lines step aside. `wrapTight` and `wrapThrough` are read as this too. */
    | 'square'
    /** The picture sits over or under the text and nothing moves. */
    | 'none';

/** What a float's position is measured from. */
export type RelativeTo = 'column' | 'margin' | 'page' | 'paragraph';

export interface FloatPosition {
    readonly relativeTo: RelativeTo;
    /** Distance from that origin. Ignored when {@link align} is set. */
    readonly offsetPx: number;
    /** `left`/`center`/`right` horizontally, or `top`/`bottom` vertically. */
    readonly align?: string;
}

/** Where a float sits and how text behaves near it — true of every kind. */
export interface FloatFrame {
    readonly horizontal: FloatPosition;
    readonly vertical: FloatPosition;
    readonly wrap: WrapMode;
    /** `distT`/`distB`/`distL`/`distR` — the clear space text keeps from it. */
    readonly marginTopPx: number;
    readonly marginBottomPx: number;
    readonly marginLeftPx: number;
    readonly marginRightPx: number;
    /** `behindDoc` — drawn under the text rather than over it. */
    readonly behindText: boolean;
}

export interface FloatingImage extends FloatFrame {
    readonly image: InlineImage;
}

/**
 * A text box: a little page of its own, anchored like a picture.
 *
 * Its size is on the FRAME rather than taken from what is inside it, because
 * that is the way round the file says it — a box states its extent, and the
 * text is broken to fit. Measured against LibreOffice: the same
 * fourteen words in a 120pt box came out five to a line with the default inset
 * and six with the inset set to zero, so the width that breaks the text is the
 * frame's less its own inset.
 */
export interface FloatingBox extends FloatFrame {
    /**
     * BLOCKS rather than paragraphs: a box can hold a table, and holding one
     * is the ordinary case for a sidebar.
     */
    readonly blocks: readonly Block[];
    readonly widthPx: number;
    readonly heightPx: number;
    readonly inset: BoxInset;
}

/**
 * The clear space between a box's frame and the words inside it.
 *
 * Four sides because that is the shape the format states, though only three do
 * work here: the left and right come off the width the text is broken at, the
 * top pushes the first line down, and the bottom would matter only to a box
 * that clipped its text or grew to fit it — neither of which is measured.
 */
export interface BoxInset {
    readonly leftPx: number;
    readonly topPx: number;
    readonly rightPx: number;
    readonly bottomPx: number;
}

/** Anything anchored beside the text: a picture, or a box of words. */
export type Float = FloatingImage | FloatingBox;

export function isFloatingBox(float: Float): float is FloatingBox {
    return 'blocks' in float;
}

/** A float once it has been given a place on a page. */
export interface PlacedFloat {
    /** A picture float. Absent on a text box, which carries lines instead. */
    readonly image?: InlineImage;
    /**
     * A text box's own lines, already moved to where the box sits on the page.
     *
     * Lines rather than blocks because the box is laid out as it is placed:
     * its width is known then, and nothing downstream re-breaks it.
     */
    readonly lines?: readonly PlacedLine[];
    /**
     * A text box's own TABLES, moved the same way its lines are.
     *
     * A box was dropping them outright — `PlacedFloat` carried lines, and
     * rows wanted the renderer's row path as well as its line path, which
     * is one call. Measured: a 2x2 table in a box at 180pt drew
     * its cells at 261.25 and 321.25 and its rules from 255.30 to 376.30.
     */
    readonly rows?: readonly PlacedRow[];
    readonly xPx: number;
    readonly yPx: number;
    readonly widthPx: number;
    readonly heightPx: number;
    readonly wrap: WrapMode;
    readonly behindText: boolean;
    /** The space text keeps clear, already added to the picture's own box. */
    readonly exclusion: Rect;
}

export interface Rect {
    readonly xPx: number;
    readonly yPx: number;
    readonly widthPx: number;
    readonly heightPx: number;
}

/** The horizontal room left for a line, once the floats beside it are out of it. */
export interface LineBox {
    readonly leftPx: number;
    readonly widthPx: number;
}

/**
 * What is left of a column for a line at a given height.
 *
 * ## Text takes the WIDER side, and only one side
 *
 * Word can split a line into runs either side of a float. This does not: a line
 * overlapping a float is given whichever side has more room, and the other is
 * left empty. Measured against LibreOffice, which put every wrapped line to the
 * RIGHT of a float 36pt from the margin — the 27pt left of it could not hold a
 * word — and the same rule gives the same answer there.
 *
 * `none`-wrapped floats are ignored entirely: they sit over or under the text
 * and nothing steps aside.
 */
export function lineBoxAt(
    column: LineBox,
    yPx: number,
    heightPx: number,
    floats: readonly PlacedFloat[],
): LineBox {
    let box = column;

    for (const float of floats) {
        if ('none' === float.wrap || !overlapsVertically(float.exclusion, yPx, heightPx)) {
            continue;
        }

        const left = float.exclusion.xPx;
        const right = float.exclusion.xPx + float.exclusion.widthPx;
        const roomLeft = left - box.leftPx;
        const roomRight = box.leftPx + box.widthPx - right;

        if (roomLeft <= 0 && roomRight <= 0) {
            // The float spans the column: nothing fits beside it, and a line
            // given a negative width would break after every character.
            return { leftPx: box.leftPx, widthPx: 0 };
        }

        box = roomRight >= roomLeft
            ? { leftPx: right, widthPx: roomRight }
            : { leftPx: box.leftPx, widthPx: roomLeft };
    }

    return box;
}

/**
 * Whether a line of this height starting here meets the float's exclusion.
 *
 * Half-open at both ends: a line whose top is exactly the float's bottom is
 * clear of it, which is the line LibreOffice gives back the full column.
 */
function overlapsVertically(exclusion: Rect, yPx: number, heightPx: number): boolean {
    return yPx < exclusion.yPx + exclusion.heightPx && yPx + heightPx > exclusion.yPx;
}
