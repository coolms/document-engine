import type { TrueTypeFont } from '../font/truetype-font.js';
import { advanceOf, opensBorderSpan, type LinePiece }
    from '../layout/line-breaker.js';
import type { BorderSide, BorderStyle, PlacedParagraphBorder }
    from '../layout/borders.js';
import type { PlacedFloat } from '../layout/float.js';
import type { InlineImage } from '../layout/image.js';
import type { Page, PlacedFurniture, PlacedLine, PlacedRow } from '../layout/page-layout.js';
import { translateLine, translateRow } from '../layout/table-layout.js';

/**
 * Turning a laid-out page into positioned draw operations.
 *
 * The last step before ink. Everything above this decided WHERE things go; this
 * flattens it into a list a renderer can walk without knowing what a section, a
 * table or a header is. The SVG serialiser beside it is one consumer; a canvas
 * or DOM renderer would be another, and they must not each re-derive geometry.
 *
 * ## Baselines, not boxes
 *
 * Every operation is positioned at its BASELINE, because that is where a
 * renderer draws text. Emitting line boxes would make each consumer work the
 * baseline out again, and they would not all work it out the same way.
 *
 * ## The page must be enough
 *
 * `renderPage` takes a page and nothing else. That constraint is the point: it
 * is what forced the baseline, the list marker and the piece's own font onto
 * the layout instead of leaving a renderer to fetch them from the document. If
 * something cannot be drawn from the page alone, the layout is missing it.
 */

export interface TextOp {
    readonly kind: 'text';
    readonly xPx: number;
    /** The BASELINE, not the top of the line. */
    readonly yPx: number;
    /**
     * A quarter turn about this operation's own origin, for `w:textDirection`.
     *
     * `ccw` reads bottom-to-top and `cw` top-to-bottom; measured off the PDF's
     * text matrix, LibreOffice writes them as `[0 1 -1 0]` and `[0 -1 1 0]`.
     * Absent means upright, which is every operation this engine has emitted
     * until now — the field is here so a renderer can turn one, and the layout
     * that decides WHICH is a second measuring pass, still to come.
     */
    readonly turn?: 'ccw' | 'cw';
    readonly text: string;
    readonly font: TrueTypeFont;
    readonly sizePx: number;
    /** `#RRGGBB`, or absent to let the renderer choose. */
    readonly colorHex?: string;
    /**
     * Whether to KERN this string, from the run's `w:kern`.
     *
     * A renderer must be told, because it draws the text as a string and the
     * drawing engine steps between the glyphs itself — kerning as it sees fit,
     * which for every browser means kerning by default. This engine measured
     * the string unkerned unless the document asked, so an operation
     * drawn on the other rule lands its later glyphs off the measurement.
     */
    readonly kerned?: boolean;
}

/** A filled rectangle — cell shading, and the paper itself. */
export interface RectOp {
    readonly kind: 'rect';
    readonly xPx: number;
    readonly yPx: number;
    readonly widthPx: number;
    readonly heightPx: number;
    readonly fill: string;
}

/**
 * A straight rule — one edge of one cell.
 *
 * Given as a line rather than a thin rectangle because a border has a JOIN and
 * a cap: two edges meeting at a corner want to be strokes, and a renderer that
 * received rectangles would have to reconstruct that.
 */
export interface LineOp {
    readonly kind: 'line';
    readonly x1Px: number;
    readonly y1Px: number;
    readonly x2Px: number;
    readonly y2Px: number;
    readonly widthPx: number;
    readonly color: string;
    readonly style: BorderStyle;
}

/**
 * A picture, placed by its TOP-LEFT corner.
 *
 * By its top-left because that is what every drawing surface takes, even though
 * the layout knows it by its baseline — converting once here beats every
 * consumer subtracting the height and one of them forgetting.
 */
export interface ImageOp {
    readonly kind: 'image';
    readonly xPx: number;
    readonly yPx: number;
    readonly widthPx: number;
    readonly heightPx: number;
    readonly image: InlineImage;
}

export type DrawOp = TextOp | RectOp | LineOp | ImageOp;

export interface RenderedPage {
    readonly widthPx: number;
    readonly heightPx: number;
    /** In paint order: furniture first, then the body over it. */
    readonly ops: readonly DrawOp[];
    readonly pageNumber: number;
}

export function renderPage(page: Page): RenderedPage {
    const ops: DrawOp[] = [];

    const furniture = (placed: PlacedFurniture | undefined): void => {
        pushRows(ops, placed?.rows ?? []);
        for (const box of placed?.paragraphBorders ?? []) {
            pushParagraphBox(ops, box);
        }
        for (const line of placed?.lines ?? []) {
            pushLine(ops, line);
        }
    };

    // Furniture first, so body text drawn over it wins any overlap — which is
    // what happens when a header outgrows the margin it was measured for.
    furniture(page.header);
    furniture(page.footer);

    // A float declaring `behindDoc` goes UNDER the text; the rest go over it,
    // after everything else has been drawn.
    for (const float of page.floats) {
        if (float.behindText) {
            pushFloat(ops, float);
        }
    }

    // Shading, then rules, then the text inside them — a cell's fill drawn
    // after its text would paint over the words.
    pushRows(ops, page.rows);

    for (const line of page.lines) {
        pushLine(ops, line);
    }

    // The numbers down the margin, each already right-aligned where it stands.
    for (const number of page.lineNumbers) {
        ops.push({
            kind: 'text',
            text: number.run.text,
            xPx: number.xPx,
            yPx: number.baselinePx,
            font: number.run.font,
            sizePx: number.run.sizePx,
            ...(undefined === number.run.colorHex ? {} : { colorHex: number.run.colorHex }),
            ...(true === number.run.kerned ? { kerned: true } : {}),
        });
    }

    // The page's own border first: it is furniture, and a paragraph that draws
    // a box of its own draws it over the top rather than under.
    for (const box of [
        ...(undefined === page.pageBorder ? [] : [page.pageBorder]),
        ...page.paragraphBorders,
    ]) {
        pushParagraphBox(ops, box);
    }

    // The notes at the foot, under the rule Word draws above them.
    if (undefined !== page.footnotes) {
        const { separatorYPx, separatorLeftPx, separatorWidthPx } = page.footnotes;
        ops.push({
            kind: 'line',
            x1Px: separatorLeftPx,
            y1Px: separatorYPx,
            x2Px: separatorLeftPx + separatorWidthPx,
            y2Px: separatorYPx,
            // A HAIRLINE, which is the thinnest rule LibreOffice will draw and
            // what it draws this one at: measured at 0.100 against
            // the whole point this used to be — seven and a half times too
            // heavy for a rule whose whole job is to be unobtrusive. Its
            // place, its length and its height above the notes were all
            // measured long ago; its thickness never was.
            widthPx: MINIMUM_RULE_PX,
            color: '#000000',
            style: 'solid',
        });
        for (const line of page.footnotes.lines) {
            pushLine(ops, line);
        }
    }

    for (const float of page.floats) {
        if (!float.behindText) {
            pushFloat(ops, float);
        }
    }

    return {
        widthPx: page.geometry.widthPx,
        heightPx: page.geometry.heightPx,
        pageNumber: page.pageNumber,
        ops,
    };
}

/**
 * A float, drawn at the place the layout gave it.
 *
 * By its top-left already — unlike an inline picture, a float never sat on a
 * baseline, so there is nothing to convert.
 *
 * A TEXT BOX arrives here as lines instead of a picture, already at page
 * coordinates, and is drawn by the same `pushLine` the body uses. That is the
 * whole of the difference: a box's text is ordinary text that was broken to a
 * narrower width.
 */
function pushFloat(ops: DrawOp[], float: PlacedFloat): void {
    if (undefined !== float.image) {
        ops.push({
            kind: 'image',
            xPx: float.xPx,
            yPx: float.yPx,
            widthPx: float.widthPx,
            heightPx: float.heightPx,
            image: float.image,
        });
    }

    // Rows before lines, as everywhere else: a cell's fill must not cover
    // the text standing in it.
    pushRows(ops, float.rows ?? []);
    for (const line of float.lines ?? []) {
        pushLine(ops, line);
    }
}

/**
 * A table's rows, and any table nested inside one of their cells.
 *
 * Shading first for every row, then every border, then the text: a cell's fill
 * must not cover the rule its neighbour draws. A nested table is painted after
 * the cell that holds it, which is the order it sits in anyway.
 */
function pushRows(ops: DrawOp[], rows: readonly PlacedRow[]): void {
    for (const row of rows) {
        pushShading(ops, row);
    }
    rows.forEach((row, index) => {
        // A shared edge is ONE rule, not two. Every row draws its own
        // bottom, and only the first of its table on this page draws a
        // top — measured against LibreOffice, which printed 63 rules for
        // 62 rows on the first page of a split table and 9 for 8 on the
        // second: one per row, plus a top on each part.
        const opens = 0 === index || rows[index - 1]?.blockIndex !== row.blockIndex;
        pushBorders(ops, row, opens);
    });
    for (const row of rows) {
        for (const cell of row.cells) {
            // The cell's own paragraph boxes, under its text the same way
            // the page's are under the body's.
            for (const box of cell.paragraphBorders) {
                pushParagraphBox(ops, box);
            }
            for (const line of cell.lines) {
                if (undefined === cell.turn) {
                    pushLine(ops, line);
                } else {
                    pushTurnedLine(ops, line, cell.turn);
                }
            }
            pushRows(ops, cell.rows);
        }
    }
}

function pushShading(ops: DrawOp[], row: PlacedRow): void {
    for (const cell of row.cells) {
        if (undefined === cell.shadingFill) {
            continue;
        }

        ops.push({
            kind: 'rect',
            xPx: cell.xPx,
            yPx: row.yPx,
            widthPx: cell.widthPx,
            heightPx: cell.heightPx,
            fill: cell.shadingFill,
        });
    }
}

/**
 * The four rules around each cell.
 *
 * Drawn down the CENTRE of the cell's edge, which is where a stroke of a given
 * width sits: offsetting by half the width instead would put a two-point border
 * entirely inside one cell and leave a gap against its neighbour.
 */
function pushBorders(ops: DrawOp[], row: PlacedRow, opensHere: boolean): void {
    for (const cell of row.cells) {
        const borders = cell.borders;
        if (undefined === borders) {
            continue;
        }

        const left = cell.xPx;
        const right = cell.xPx + cell.widthPx;
        // A horizontal rule is CENTRED in the gap the flow kept for it, so it
        // stands half its width OUTSIDE the content it belongs to — measured
        // against LibreOffice, a row of 11.5pt text between one-point rules
        // ran 800.97 to 789.47 with its rules at 801.489 and 788.989.
        const halfTop = (borders.top?.widthPx ?? 0) / 2;
        const halfBottom = (borders.bottom?.widthPx ?? 0) / 2;
        const top = row.yPx - halfTop;
        // The cell's own box, not the row's: a vertically merged cell is drawn
        // with the row it starts in and reaches past the foot of it.
        const bottom = row.yPx + cell.heightPx + halfBottom;

        if (opensHere) {
            rule(ops, borders.top, left, top, right, top);
        }
        rule(ops, borders.bottom, left, bottom, right, bottom);
        // The verticals run half a horizontal rule PAST it at each end, rather
        // than stopping on its centre line, which is what fills the corner in.
        // Measured: a table between one-point rules at 769.389 and
        // 719.189 printed its left edge from 769.889 to 718.689 — half a rule
        // beyond each. Stopping on the line leaves a quarter of every outer
        // corner unpainted.
        //
        // Half of THIS cell's rule, which is right wherever it shows: at an
        // inside junction the cell's own top edge has been resolved away — the
        // row above draws it — and the gap that leaves is underneath the
        // horizontal rule itself, which is a whole width tall.
        rule(ops, borders.left, left, top - halfTop, left, bottom + halfBottom);
        rule(ops, borders.right, right, top - halfTop, right, bottom + halfBottom);
    }
}

function rule(
    ops: DrawOp[],
    side: BorderSide | undefined,
    x1Px: number,
    y1Px: number,
    x2Px: number,
    y2Px: number,
): void {
    if (undefined === side) {
        return;
    }

    ops.push({
        kind: 'line',
        x1Px,
        y1Px,
        x2Px,
        y2Px,
        widthPx: side.widthPx,
        color: side.colorHex,
        style: side.style,
    });
}

/**
 * Where one of a decoration's rules goes, given how far from the baseline it
 * sits — positive below, negative above.
 *
 * Upright and turned text disagree only about which direction that is, so WHAT
 * a decoration is (an underline the font's offset below the baseline, a second
 * rule a thickness below that, a strike above) is decided in one place and
 * placed in two.
 */
type PushRule = (
    offsetPx: number,
    thickness: number,
    style: BorderStyle,
    stroke: string,
) => void;

/**
 * The mark a highlighted run leaves, if it has one.
 *
 * The box is the LINE, not the font: verified against LibreOffice, where a run
 * highlighted inside an exactly-spaced 24pt line fills all 24 points rather
 * than the 11.5 the font would occupy. Filling the font's box instead leaves a
 * pale stripe above and below every highlight in a loosely spaced document.
 */
function pushHighlight(ops: DrawOp[], piece: LinePiece, x: number, placed: PlacedLine): void {
    if (undefined === piece.highlightHex) {
        return;
    }

    ops.push({
        kind: 'rect',
        xPx: x,
        yPx: placed.yPx,
        widthPx: piece.widthPx,
        heightPx: placed.heightPx,
        fill: piece.highlightHex,
    });
}

function pushLine(ops: DrawOp[], placed: PlacedLine): void {
    const lineBaseline = placed.yPx + placed.baselinePx;
    const marker = placed.marker;

    if (undefined !== marker) {
        ops.push({
            kind: 'text',
            xPx: marker.xPx,
            yPx: lineBaseline,
            text: marker.run.text,
            font: marker.run.font,
            sizePx: marker.run.sizePx,
            ...(undefined === marker.run.colorHex ? {} : { colorHex: marker.run.colorHex }),
            ...(true === marker.run.kerned ? { kerned: true } : {}),
        });
    }

    let x = placed.xPx;
    // Where the open boxed run's text began on this line, if one is open.
    let spanStartPx: number | null = null;
    const pieces = placed.line.pieces;

    for (const [index, piece] of pieces.entries()) {
        if ('' === piece.text) {
            continue;
        }

        // The room a boxed run keeps clear is the RUN's, charged once where
        // its span opens on this line and once where it closes — so the box
        // goes round the whole span and not round each word of it.
        if (opensBorderSpan(piece, pieces[index - 1])) {
            x += piece.borderRoomPx ?? 0;
            spanStartPx = x;
        }

        // A script run draws off the line it belongs to; everything about
        // where the line SITS is unchanged.
        const baseline = lineBaseline + (piece.baselineShiftPx ?? 0);

        // Behind the text, and only behind THIS piece: adjacent pieces abut
        // rather than overlap, so painting each as it goes is safe.
        pushHighlight(ops, piece, x, placed);

        const from = x;
        // The gloss of a `w:ruby`, on a baseline of its own at the TOP of the
        // line. Measured: a 5pt gloss printed 4.68 below the line's
        // top — its own ascent — and centred over its base, which is what
        // `rubyAlign` centre asks for and what the piece's width already
        // reserves where the gloss is the wider of the two.
        if (undefined !== piece.ruby) {
            const glossPx = advanceOf(piece.ruby.text, piece.ruby);
            const font = piece.ruby.font;

            ops.push({
                kind: 'text',
                xPx: x + (piece.widthPx - glossPx) / 2,
                yPx: placed.yPx
                    + font.naturalLineHeight(piece.ruby.sizePx) - font.descent(piece.ruby.sizePx),
                text: piece.ruby.text,
                font,
                sizePx: piece.ruby.sizePx,
                ...(undefined === piece.ruby.colorHex ? {} : { colorHex: piece.ruby.colorHex }),
                ...(true === piece.ruby.kerned ? { kerned: true } : {}),
            });
        }

        // A base narrower than its own gloss is centred under it, and the
        // piece still advances by the whole group: measured, an `xy` base
        // under a gloss 28.90 wide printed at 95.45 where the group began at
        // 86.00, and the text after the group at 114.85.
        const insetPx = undefined === piece.ruby
            ? 0
            : Math.max(0, (piece.widthPx - advanceOf(piece.text, piece)) / 2);

        x = placed.wordSpacingPx > 0
            ? pushStretched(ops, piece, x + insetPx, baseline, placed.wordSpacingPx)
            : pushWhole(ops, piece, x + insetPx, baseline, index === pieces.length - 1);
        if (insetPx > 0) {
            x = from + piece.widthPx;
        }

        // Over the text, and across the whole piece — a justified line's gaps
        // are underlined too, which is why this uses the advanced width rather
        // than the piece's own.
        if (undefined !== piece.underline || undefined !== piece.strike) {
            pushDecoration(piece, piece.colorHex ?? '#000000',
                (offsetPx, thickness, style, stroke) => ops.push({
                    kind: 'line',
                    x1Px: from,
                    y1Px: baseline + offsetPx,
                    x2Px: x,
                    y2Px: baseline + offsetPx,
                    widthPx: Math.max(thickness, MINIMUM_RULE_PX),
                    color: stroke,
                    style,
                }));
        }

        // Closed where the span ends, which on a broken run is the end of the
        // line: a run over three lines came out of LibreOffice as three
        // COMPLETE boxes, each with its own top and bottom.
        if (null !== spanStartPx
            && undefined !== piece.border
            && piece.runIndex !== pieces[index + 1]?.runIndex) {
            pushRunBorder(ops, piece, piece.border, baseline, spanStartPx, x);
            x += piece.borderRoomPx ?? 0;
            spanStartPx = null;
        }
    }
}

/**
 * The box round one line's worth of a boxed run.
 *
 * The rules stand `w:space` plus half the width outside the text — a
 * space-nought box measured 19.25pt round a run whose own advance is 18.30 —
 * and run half the crossing rule past each corner, as a paragraph's box does.
 *
 * Measured off its OWN baseline and its OWN metrics, not off the line it sits
 * on: a boxed 10pt run beside a 20pt one drew exactly the same 18.45pt box, in
 * exactly the same place relative to its baseline, as one with a 10pt
 * neighbour. Taking the line's box would have grown it with the neighbour.
 */
function pushRunBorder(
    ops: DrawOp[],
    piece: LinePiece,
    border: BorderSide,
    baselinePx: number,
    fromPx: number,
    toPx: number,
): void {
    const standoff = (border.spacePx ?? 0) + border.widthPx / 2;
    const half = border.widthPx / 2;

    const descentPx = piece.font.descent(piece.sizePx);
    const topPx = baselinePx - (piece.font.naturalLineHeight(piece.sizePx) - descentPx)
        - standoff;
    const bottomPx = baselinePx + descentPx + standoff;
    const leftPx = fromPx - standoff;
    const rightPx = toPx + standoff;

    const rule = (x1Px: number, y1Px: number, x2Px: number, y2Px: number): void => {
        ops.push({
            kind: 'line',
            x1Px,
            y1Px,
            x2Px,
            y2Px,
            widthPx: border.widthPx,
            color: border.colorHex,
            style: border.style,
        });
    };

    rule(leftPx - half, topPx, rightPx + half, topPx);
    rule(leftPx - half, bottomPx, rightPx + half, bottomPx);
    rule(leftPx, topPx - half, leftPx, bottomPx + half);
    rule(rightPx, topPx - half, rightPx, bottomPx + half);
}

/**
 * What each leader is made of.
 *
 * `heavy` is Word's thick underscore, drawn here as an ordinary one — the right
 * character at the wrong weight, which is closer than leaving the gap blank.
 */
/** A piece that is nothing but tabs — the breaker gives each its own. */
const TABS_ONLY = /^\t+$/u;

const LEADER_GLYPHS: Record<string, string> = {
    dot: '.',
    hyphen: '-',
    underscore: '_',
    middleDot: '\u00b7',
    heavy: '_',
};

/**
 * How many leader glyphs fill a span `exact` of them long.
 *
 * Three answers for five leaders, each measured off a printed page
 * rather than reasoned about — the spans are in `tab-leader-fill.docx`, chosen
 * to land just over, halfway, just under and exactly on a glyph, because only
 * the exact case tells `ceil` from "one more than fits".
 *
 * There is no theory here that unifies them. A dot leader never reaches its
 * stop, a rule always passes it, and the two middling ones go to whichever is
 * nearer; that is what the page does, and the page is the reference.
 */
function leaderCount(leader: string, exact: number): number {
    switch (leader) {
        // One MORE than fits, so a rule to sign on always reaches its stop.
        // 36.000 printed 37, which is what rules out `ceil`; 36.100, 36.500
        // and 36.900 printed 37; 35.991 printed 36; and 0.500 printed one
        // single underscore, overshooting a span half its own width.
        //
        // `heavy` draws the same glyph and its ONE measured span agreed.
        case 'underscore':
        case 'heavy':
            return Math.floor(exact) + 1;
        // Whichever is nearer: 54.804 printed 55, 54.053 printed 54. But
        // nothing at all where a whole glyph does not fit — 0.546 printed
        // none, which rounding alone would have made one.
        case 'hyphen':
        case 'middleDot':
            return exact < 1 ? 0 : Math.round(exact);
        // Dots never overshoot. 140.109, 140.509, 140.909 and 140.982 all
        // printed 140 — so not `round`, which would have given 141 for three
        // of them — and an exact 141.000 printed 141.
        default:
            return Math.floor(exact);
    }
}

/**
 * Fill a tab's span with its leader.
 *
 * Starting where the tab does — LibreOffice begins its dots immediately after
 * the text before the tab, not ranged against the stop — and as many glyphs as
 * {@link leaderCount} says, which is not always as many as fit.
 */
function pushLeader(ops: DrawOp[], piece: LinePiece, x: number, baseline: number): void {
    const glyph = LEADER_GLYPHS[piece.leader ?? ''];
    if (undefined === glyph) {
        return;
    }

    const advance = piece.font.measureAdvance(glyph, piece.sizePx).widthPt;
    const count = advance > 0 ? leaderCount(piece.leader ?? '', piece.widthPx / advance) : 0;
    if (count <= 0) {
        return;
    }

    ops.push({
        kind: 'text',
        xPx: x,
        yPx: baseline,
        text: glyph.repeat(count),
        font: piece.font,
        sizePx: piece.sizePx,
        ...(undefined === piece.colorHex ? {} : { colorHex: piece.colorHex }),
        ...(true === piece.kerned ? { kerned: true } : {}),
    });
}

/**
 * The rules under and through a piece of text.
 *
 * Positioned from the FONT's own metrics rather than a constant: a rule a fixed
 * distance below every baseline sits on the descenders of one face and floats
 * away from another.
 *
 * Drawn after the text so a heavy rule is not hidden by it, and given the
 * text's own colour unless the underline named one.
 */
function pushDecoration(piece: LinePiece, colour: string, rule: PushRule): void {
    const under = piece.underline;
    if (undefined !== under) {
        const offset = piece.font.underlineOffset(piece.sizePx);
        const thickness = piece.font.underlineThickness(piece.sizePx);
        const stroke = under.colorHex ?? colour;
        const style: BorderStyle = 'dotted' === under.style || 'dashed' === under.style
            ? under.style
            : 'solid';

        rule(offset, thickness, style, stroke);
        if ('double' === under.style) {
            // The second rule sits a whole thickness below the first, so the
            // pair reads as two lines rather than one heavy one.
            //
            // MEASURED, NOT MATCHED. LibreOffice draws a double
            // underline THINNER than a single one and straddles the single's
            // place with it: Liberation Serif at 10pt printed −0.80 and −2.10
            // at 0.40 thick where a single is −1.20 at 0.60, and at 40pt
            // −3.00 and −7.20 at 1.40 against −4.40 at 2.20. This engine takes
            // its single from the FONT rather than from LibreOffice — the
            // decision in `TrueTypeFont.decoration`, on six measurements — and
            // a double built on somebody else's single would be neither.
            rule(offset + thickness * 2, thickness, style, stroke);
        }
    }

    if (undefined !== piece.strike) {
        const offset = piece.font.strikeoutOffset(piece.sizePx);
        const thickness = piece.font.strikeoutThickness(piece.sizePx);

        rule(-offset, thickness, 'solid', colour);
        if ('double' === piece.strike) {
            // The same measured-not-matched divergence as the double underline
            // above: the print struck 10pt text at +3.00 and +1.70 with a 0.40
            // rule where a single is +2.60 at 0.60, and 40pt at +11.80 and
            // +7.60 with 1.40 against +10.40 at 2.20.
            rule(-offset - thickness * 2, thickness, 'solid', colour);
        }
    }
}

/**
 * The thinnest rule LibreOffice will draw: a tenth of a point.
 *
 * Measured, where this used to be an invented 0.5px. An underline is about a
 * twentieth of the font's size — 0.1, 0.2, 0.5 and 2.1pt at 2, 4, 10 and 40pt
 * — and at 1pt, where a twentieth would be 0.05, LibreOffice still drew 0.1.
 * That is the floor, and 0.5px is 0.375pt: nearly three times too thick.
 */
const MINIMUM_RULE_PX = 0.1 * 96 / 72;

/**
 * One paragraph's box — or a page's, which is drawn the same way.
 *
 * Shared because a cell draws these too: `w:pBdr` inside a table cell came out
 * of LibreOffice with the same geometry it has on the page, and a second copy
 * of this arithmetic is a second place for it to drift.
 */
function pushParagraphBox(ops: DrawOp[], box: PlacedParagraphBorder): void {
    const rule = (
        side: BorderSide | undefined,
        x1Px: number,
        y1Px: number,
        x2Px: number,
        y2Px: number,
    ): void => {
        if (undefined === side) {
            return;
        }

        ops.push({
            kind: 'line',
            x1Px,
            y1Px,
            x2Px,
            y2Px,
            widthPx: side.widthPx,
            color: side.colorHex,
            style: side.style,
        });
    };

    // Every rule runs half the CROSSING rule's width past each end, so the
    // corners close rather than leaving a notch the width of the stroke.
    // Measured off LibreOffice: a one-point box's top ran from 27.4 to 567.95
    // while its sides stood at 27.9 and 567.45.
    const half = (side: BorderSide | undefined): number =>
        undefined === side ? 0 : side.widthPx / 2;

    const left = box.leftPx - half(box.borders.left);
    const right = box.rightPx + half(box.borders.right);
    const top = box.topPx - half(box.borders.top);
    const bottom = box.bottomPx + half(box.borders.bottom);

    if (box.opensHere) {
        rule(box.borders.top, left, box.topPx, right, box.topPx);
    }
    if (box.closesHere) {
        rule(box.borders.bottom, left, box.bottomPx, right, box.bottomPx);
    }
    rule(box.borders.left, box.leftPx, top, box.leftPx, bottom);
    rule(box.borders.right, box.rightPx, top, box.rightPx, bottom);

    // `w:between`, which runs the full width of the box like its top does.
    for (const y of box.innerYPx) {
        rule(box.borders.insideH, left, y, right, y);
    }
}

/**
 * One line of a cell whose text is turned a quarter turn.
 *
 * Its own advance runs UP the page for `ccw` and down for `cw`, so the pieces
 * step along Y where an upright line steps along X, and every operation
 * carries the turn so a renderer spins it about its own origin.
 *
 * A picture inside a turned cell is still not drawn: `pushWhole` places one
 * against an upright baseline, and no measurement of a turned one exists.
 */
function pushTurnedLine(ops: DrawOp[], placed: PlacedLine, turn: 'ccw' | 'cw'): void {
    // The line's baseline runs across its own advance, so it offsets X —
    // from the box's left edge for `ccw`, whose glyph tops face left, and from
    // its right edge for `cw`, where they face the other way.
    const baselineX = 'ccw' === turn
        ? placed.xPx + placed.baselinePx
        : placed.xPx + placed.heightPx - placed.baselinePx;
    const forward = 'ccw' === turn ? -1 : 1;
    let along = placed.yPx;

    for (const piece of placed.line.pieces) {
        if ('' === piece.text) {
            continue;
        }

        const ends = along + forward * piece.widthPx;

        // A quarter turn leaves a rectangle axis-aligned, so a turned
        // highlight is the upright box with its sides swapped. Measured: a
        // 12.2pt line highlighted over a 14.25pt run printed as
        // `re 33.600 763.739 12.150 14.250` — the LINE across, the RUN along.
        if (undefined !== piece.highlightHex) {
            ops.push({
                kind: 'rect',
                xPx: placed.xPx,
                yPx: Math.min(along, ends),
                widthPx: placed.heightPx,
                heightPx: piece.widthPx,
                fill: piece.highlightHex,
            });
        }

        ops.push({
            kind: 'text',
            turn,
            text: piece.text,
            xPx: baselineX,
            yPx: along,
            font: piece.font,
            sizePx: piece.sizePx,
            ...(undefined === piece.colorHex ? {} : { colorHex: piece.colorHex }),
            ...(true === piece.kerned ? { kerned: true } : {}),
        ...(true === piece.kerned ? { kerned: true } : {}),
        });

        if (undefined !== piece.underline || undefined !== piece.strike) {
            pushDecoration(piece, piece.colorHex ?? '#000000',
                (offsetPx, thickness, style, stroke) => {
                    // Below the baseline becomes BESIDE it. LibreOffice draws
                    // this one inside the rotation — `0 -1.1 l 14.2 -1.1`, the
                    // run's own length at the font's offset — so the same
                    // number lands on X here, on the descender side.
                    const acrossPx = baselineX - forward * offsetPx;

                    ops.push({
                        kind: 'line',
                        x1Px: acrossPx,
                        y1Px: along,
                        x2Px: acrossPx,
                        y2Px: ends,
                        widthPx: Math.max(thickness, MINIMUM_RULE_PX),
                        color: stroke,
                        style,
                    });
                });
        }

        along = ends;
    }
}

function pushWhole(
    ops: DrawOp[],
    piece: LinePiece,
    x: number,
    baseline: number,
    isLast = false,
): number {
    if (undefined !== piece.image) {
        ops.push({
            kind: 'image',
            xPx: x,
            // Standing on the baseline: the top is a whole height above it.
            yPx: baseline - piece.image.heightPx,
            widthPx: piece.image.widthPx,
            heightPx: piece.image.heightPx,
            image: piece.image,
        });

        return x + piece.widthPx;
    }

    if (undefined !== piece.shape) {
        // Standing on the baseline like a picture: a rule or a panel sits in
        // the line rather than over it.
        const top = baseline - piece.shape.heightPx;

        // A shape may stand a little INTO the room it keeps — a VML one begins
        // a wrap distance in. Its words are already offset by that
        // and by the box's inset, stacked when it was read, so all
        // that is left is to put box and words where the box stands.
        const left = x + (piece.shape.leadPx ?? 0);

        pushRows(ops, (piece.shape.rows ?? []).map((row) => translateRow(row, x, top)));
        for (const line of piece.shape.lines ?? []) {
            pushLine(ops, translateLine(line, x, top));
        }
        if (undefined !== piece.shape.fillHex) {
            ops.push({
                kind: 'rect',
                xPx: left,
                yPx: top,
                widthPx: piece.shape.widthPx,
                heightPx: piece.shape.heightPx,
                fill: piece.shape.fillHex,
            });
        }
        if (undefined !== piece.shape.strokeHex && piece.shape.strokeWidthPx > 0) {
            const right = left + piece.shape.widthPx;
            const foot = top + piece.shape.heightPx;
            const edge = (x1Px: number, y1Px: number, x2Px: number, y2Px: number): void => {
                ops.push({
                    kind: 'line',
                    x1Px,
                    y1Px,
                    x2Px,
                    y2Px,
                    widthPx: piece.shape!.strokeWidthPx,
                    color: piece.shape!.strokeHex!,
                    style: 'solid',
                });
            };

            edge(left, top, right, top);
            edge(left, foot, right, foot);
            edge(left, top, left, foot);
            edge(right, top, right, foot);
        }

        return x + piece.widthPx;
    }

    if (undefined !== piece.leader) {
        pushLeader(ops, piece, x, baseline);

        return x + piece.widthPx;
    }

    // A tab is a distance, not a mark. Its advance is already in the positions
    // of everything after it, and drawing the character as well adds a space's
    // worth of ink — a renderer that honours `xml:space` draws one, and one
    // that does not collapses it to nothing at an unpredictable width.
    if (TABS_ONLY.test(piece.text)) {
        return x + piece.widthPx;
    }

    ops.push({
        kind: 'text',
        xPx: x,
        yPx: baseline,
        text: inkOf(piece, isLast),
        font: piece.font,
        sizePx: piece.sizePx,
        ...(undefined === piece.colorHex ? {} : { colorHex: piece.colorHex }),
        ...(true === piece.kerned ? { kerned: true } : {}),
    });

    return x + piece.widthPx;
}

/**
 * The text of a piece as it is DRAWN.
 *
 * A soft hyphen is an offer of a break: no ink where the line carries on past
 * it, and a HYPHEN where the line was ended there. Measured against
 * LibreOffice: a word split by one printed `eeeeeeee` and a `-` at the end of
 * that line, with the rest beginning the next.
 *
 * Most fonts keep a hyphen glyph at U+00AD, so drawing the character as it
 * stands would put one in the middle of every word that offers a break.
 */
function inkOf(piece: LinePiece, isLast: boolean): string {
    return isLast && true === piece.hyphenAtEnd ? piece.text + '-' : piece.text;
}

/**
 * Emit a justified piece one word at a time, so each gap can be widened.
 *
 * The stretch was divided by GAPS, and a run of several spaces is one gap — so
 * it is added once per run and not once per space. Adding it per character
 * would push every line containing a double space past the margin it was fitted
 * to, and the line would no longer end where justification promised.
 *
 * Each part is measured with the piece's own font, so this is the same
 * measurement the breaker made rather than a second implementation of it. The
 * only difference is kerning across a split, which is the pair either side of a
 * space and is the same caveat the breaker already documents for run
 * boundaries.
 */
function pushStretched(
    ops: DrawOp[],
    piece: LinePiece,
    startX: number,
    baseline: number,
    wordSpacingPx: number,
): number {
    let x = startX;

    // Capturing split, so the gaps survive: a gap has to be measured, not
    // skipped, or every word after the first lands a space too far left.
    for (const part of piece.text.split(/( +)/u)) {
        if ('' === part) {
            continue;
        }

        const width = advanceOf(part, piece);

        if (' ' === part[0]) {
            x += width + wordSpacingPx;

            continue;
        }

        ops.push({
            kind: 'text',
            xPx: x,
            yPx: baseline,
            text: part,
            font: piece.font,
            sizePx: piece.sizePx,
            ...(undefined === piece.colorHex ? {} : { colorHex: piece.colorHex }),
            ...(true === piece.kerned ? { kerned: true } : {}),
        ...(true === piece.kerned ? { kerned: true } : {}),
        });
        x += width;
    }

    return x;
}
