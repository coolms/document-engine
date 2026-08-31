/**
 * Pictures.
 *
 * An inline image is not a block. Word puts it in a RUN, so it sits in the line
 * like a very large character: it takes width from the line it is on, it makes
 * that line taller, and text can follow it on the same line. Modelling it as a
 * paragraph of its own would be right for the common case of an image alone in
 * a paragraph and wrong for every icon in a sentence — and the common case
 * falls out of the general one for free.
 */

// Type-only, so nothing is imported at run time and the cycle with
// `page-layout` — which imports the shapes below — never forms.
import type { PlacedLine, PlacedRow } from './page-layout.js';

/** The bytes of a picture, and enough to hand them to a renderer. */
export interface ImageContent {
    readonly bytes: Uint8Array;
    /** `image/png`, `image/jpeg`, … taken from the package's content types. */
    readonly contentType: string;
    /** `wp:docPr/@descr` — the alternative text, when the document supplies one. */
    readonly description?: string;
}

/**
 * A picture placed in a run.
 *
 * The size is the one the DOCUMENT asks for (`wp:extent`), not the picture's
 * own: a 2000px photograph scaled into a 3-inch box is 3 inches wide, and
 * measuring the file instead would blow the layout apart.
 */
export interface InlineImage {
    readonly content: ImageContent;
    readonly widthPx: number;
    readonly heightPx: number;
    /**
     * Where this picture sits ACROSS its line, when it does not stand on the
     * baseline like a glyph.
     *
     * Absent is the ordinary case and the only one an upright line has. A
     * TURNED cell is different, and measured against LibreOffice it
     * is different in two ways, both set by `turnPictures`:
     *
     * - `centred` — the picture is centred on the line's ASCENT, so the
     *   baseline lands half way between the ascender edge and the picture's
     *   foot. Printed baselines 13.70 and 22.70 from the cell's inner edge for
     *   pictures 18 and 36 across, against 9.40 for the ascent alone.
     * - `stacked` — the text hangs entirely BELOW the picture. What LibreOffice
     *   does with a picture too long for its line is give it a line of its own;
     *   kept on the text's line, that is the same geometry.
     */
    readonly acrossLine?: 'centred' | 'stacked';
}

/**
 * A drawn box standing in a run's place — a VML shape with no picture in it.
 *
 * A rule, a border, a coloured panel: furniture rather than content, and the
 * thing a `w:pict` is when it holds no `v:imagedata`. It occupies a run the
 * same way a picture does, which is why it lives beside one.
 */
export interface InlineShape {
    readonly widthPx: number;
    readonly heightPx: number;
    /** `#RRGGBB`, or absent for a shape that is drawn but not filled. */
    readonly fillHex?: string;
    /** `#RRGGBB`, or absent for one with no outline. */
    readonly strokeHex?: string;
    readonly strokeWidthPx: number;
    /**
     * How far INTO its room the shape is drawn, where that is not nought.
     *
     * The room a VML shape keeps is wrap distance either side of it, so the
     * box itself begins a wrap distance in — 9pt by default. Left unbuilt
     * while only the room was measured; forced by the interior, where the
     * words print 13.2 from the shape's origin and that only decomposes as
     * this 9.0 plus the box's own 4.25 inset. Draw the box at the origin and
     * the same words would sit 13.2 inside a box they should sit 4.25 inside.
     */
    readonly leadPx?: number;
    /**
     * The room the shape takes in the LINE, where that differs from its size.
     *
     * A VML shape is drawn at the width it states and then keeps 18.0pt more
     * of the line than it draws — measured at 30, 60, 90, 150 and 180pt, on
     * `v:shape` and `v:rect` alike, with and without a text box inside. Two
     * numbers, so two fields: widen the one the renderer paints and the box
     * grows, which the print does not.
     *
     * Absent means the shape takes exactly the room it draws.
     */
    readonly advanceWidthPx?: number;
    /**
     * The words inside it, when the shape is a TEXT BOX sitting in the line.
     *
     * Already stacked, and offset by the box's own inset — so a renderer adds
     * the piece's origin and nothing else. They can be laid out this early
     * because an inline box's inner width is settled by its `wp:extent` before
     * anything knows where the box lands, which is not true of a float: that
     * one is stacked at placement, in `floatContent`.
     *
     * Measured: a 90x36pt box takes exactly its extent in the line and draws
     * its words 7.20 inside it.
     */
    readonly lines?: readonly PlacedLine[];
    /**
     * And its TABLES, moved the same way.
     *
     * Carried for the reason the float's are and found the way
     * that one was: a box that could hold a table met a box that carried
     * only lines, and the table went out silently between them.
     */
    readonly rows?: readonly PlacedRow[];
}

/** A single UTF-16 unit standing in for the picture, so offsets still line up. */
export const OBJECT_REPLACEMENT = '￼';
