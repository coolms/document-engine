/**
 * The unit conversions OOXML forces on anyone reading a Word document.
 *
 * ## Why this is its own module
 *
 * A single `.docx` measures itself in five different units at once: page size
 * and indents in TWIPS, font size in HALF-POINTS, border widths in
 * EIGHTH-POINTS, drawing geometry in EMU, and line spacing in twips again but
 * with a mode flag that changes what the number means. Nothing in the file says
 * which is which — `w:sz` is half-points on `w:rPr` and eighth-points on
 * `w:tcBorders`, spelled identically. Getting one wrong does not throw; it
 * produces a document that is subtly the wrong size, which is the hardest kind
 * of bug to see.
 *
 * ## The engine's internal unit is the CSS pixel
 *
 * 96 per inch, because the layout must agree with the browser that draws it and
 * that is the browser's unit. Every conversion happens here, at the OOXML
 * boundary, so no other module has to know a document was ever measured in
 * twips.
 *
 * The conversion is exact: 1pt = 96/72 = 4/3 px, and both 96 and 72 are exactly
 * representable, so points and pixels convert without loss in either direction.
 *
 * ## Divide, do not multiply by the reciprocal
 *
 * `twips * 0.05` looks identical to `twips / 20` and is not: 0.05 has no exact
 * binary representation, and the two disagree for 7186 of the first 20 000 twip
 * values. Every function here divides.
 *
 * The difference is NOT visible when writing back — the rounding below absorbs
 * it — which is exactly why the rule is easy to break and hard to notice. It is
 * visible in COMPARISONS: line breaking asks `width <= available`, and an
 * epsilon on the wrong side of that ends the line one word early.
 */

/** 1440 twips to the inch — a twip is 1/20 of a point, as OOXML defines it. */
export const TWIPS_PER_INCH = 1440;
export const POINTS_PER_INCH = 72;
/**
 * English Metric Units: 914400 to the inch, as the OOXML spec defines them —
 * chosen there to divide evenly by both 72 and 25.4.
 */
export const EMU_PER_INCH = 914400;
/** CSS reference pixel. Not a device pixel, and not the screen's actual DPI. */
export const PX_PER_INCH = 96;

const TWIPS_PER_POINT = TWIPS_PER_INCH / POINTS_PER_INCH; // 20
const EMU_PER_POINT = EMU_PER_INCH / POINTS_PER_INCH; // 12700

export function twipsToPoints(twips: number): number {
    return twips / TWIPS_PER_POINT;
}

export function twipsToPx(twips: number): number {
    return pointsToPx(twipsToPoints(twips));
}

export function pointsToPx(points: number): number {
    return (points * PX_PER_INCH) / POINTS_PER_INCH;
}

export function pxToPoints(px: number): number {
    return (px * POINTS_PER_INCH) / PX_PER_INCH;
}

/**
 * `w:sz` on a run: font size in HALF-points. `w:sz w:val="28"` is 14pt.
 *
 * Word stores it this way so half-point sizes need no decimal separator, which
 * would have been locale-dependent in the format's ancestry.
 */
export function halfPointsToPoints(halfPoints: number): number {
    return halfPoints / 2;
}

export function halfPointsToPx(halfPoints: number): number {
    return pointsToPx(halfPointsToPoints(halfPoints));
}

/**
 * `w:sz` on a BORDER: width in EIGHTH-points. The same attribute name as a font
 * size and a different unit — 8 is 1pt here and 4pt there.
 */
export function eighthPointsToPoints(eighthPoints: number): number {
    return eighthPoints / 8;
}

export function eighthPointsToPx(eighthPoints: number): number {
    return pointsToPx(eighthPointsToPoints(eighthPoints));
}

export function emuToPoints(emu: number): number {
    return emu / EMU_PER_POINT;
}

export function emuToPx(emu: number): number {
    return pointsToPx(emuToPoints(emu));
}

/**
 * Back to twips for writing.
 *
 * ROUNDS, because OOXML attribute values are integers — a round trip through
 * pixels leaves about 1.5e-11 of error, and emitting `w:w="12240.000000000002"`
 * produces a file Word refuses to open. Rounding here rather than at each call
 * site means there is one place where the format's integer requirement is
 * honoured, instead of a convention everyone is expected to remember.
 */
export function pointsToTwips(points: number): number {
    return Math.round(points * TWIPS_PER_POINT);
}

export function pxToTwips(px: number): number {
    return pointsToTwips(pxToPoints(px));
}

export function pointsToHalfPoints(points: number): number {
    return Math.round(points * 2);
}

export function pxToHalfPoints(px: number): number {
    return pointsToHalfPoints(pxToPoints(px));
}

export function pointsToEmu(points: number): number {
    return Math.round(points * EMU_PER_POINT);
}

export function pxToEmu(px: number): number {
    return pointsToEmu(pxToPoints(px));
}
