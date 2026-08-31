import type { Line } from './line-breaker.js';

/**
 * `w:jc` — where a line sits in the space available to it.
 *
 * Its own module rather than part of `page-layout`, because the page flow and
 * table cells both need it and `table-layout` imports `page-layout` for TYPES
 * ONLY: pulling a function across that boundary would make the two a runtime
 * import cycle. One shared implementation is the point — a heading centred one
 * way outside a table and another way inside it is the bug this prevents.
 */

/**
 * OOXML spells the first two `left`/`right` in Word's own files and
 * `start`/`end` in the strict schema; both are read into these. `distribute`
 * is read as `justify`, which stretches the same gaps but leaves the last line
 * alone where `distribute` would stretch it too.
 */
export type Alignment = 'left' | 'center' | 'right' | 'justify';

/**
 * Where a line sits within the width available to it, and how far its spaces
 * stretch.
 *
 * ## The last line of a justified paragraph is not justified
 *
 * Nor is one the author ended themselves with a break. Either would otherwise
 * be stretched across the whole column from a handful of words, which is the
 * most recognisable way justified text goes wrong.
 *
 * ## An overflowing line is never pushed further out
 *
 * A line wider than its column — a long URL — has negative slack. Centring or
 * right-aligning on that would move it LEFT of the indent, so the offset is
 * clamped and the line overflows to the right, as Word lets it.
 */
export function alignLine(
    line: Line,
    availableWidthPx: number,
    alignment: Alignment = 'left',
    isLastLine = false,
): { offsetPx: number; wordSpacingPx: number } {
    const slack = availableWidthPx - line.widthPx;

    if ('justify' === alignment) {
        const stretchable = !isLastLine && !line.endedByMandatoryBreak && line.spaceGaps > 0;

        return { offsetPx: 0, wordSpacingPx: stretchable ? Math.max(0, slack) / line.spaceGaps : 0 };
    }

    if ('center' === alignment) {
        return { offsetPx: Math.max(0, slack) / 2, wordSpacingPx: 0 };
    }

    if ('right' === alignment) {
        return { offsetPx: Math.max(0, slack), wordSpacingPx: 0 };
    }

    return { offsetPx: 0, wordSpacingPx: 0 };
}
