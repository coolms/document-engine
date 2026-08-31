/**
 * Where the baseline sits inside a line box.
 *
 * A line's height says how much room it takes; it does not say where the text
 * within it rests. Get this wrong and every glyph on the page is drawn at the
 * wrong height — the page count stays right and the document is visibly wrong.
 *
 * ## The three rules disagree, and only about this
 *
 * `w:spacing/@w:lineRule` picks how a line's height is arrived at, and the
 * reader resolves all three to one number because the FLOW only needs the
 * number. Drawing needs more than that: two lines of identical height put their
 * baseline in different places depending on which rule produced it.
 *
 * All three were measured out of LibreOffice's own PDF — the `Td` operand of
 * each line — at three heights each, and confirmed twice per height by solving
 * consecutive baseline-to-baseline gaps rather than trusting any assumption
 * about where a line box ended.
 */

/** `w:spacing/@w:lineRule`. Absent means `auto`, which is what Word assumes. */
export type LineRule = 'auto' | 'exact' | 'atLeast' | 'grid';

/**
 * The proportion of an EXACTLY spaced line that sits above the baseline.
 *
 * A flat proportion, not a font metric — verified by setting the same exact
 * height in Liberation Serif and Liberation Sans, whose natural baselines
 * differ by 0.05pt, and finding the exact baseline identical in both at
 * 19.18pt of a 24pt line. 0.8 × 24 is 19.2, and the 0.02 is the constant inset
 * LibreOffice applies to every line including left-aligned ones.
 *
 * So a font with a deep descender loses its descender to the line below rather
 * than moving the baseline up, which is exactly what "exact" means.
 */
export const EXACT_BASELINE_RATIO = 0.8;

export interface BaselineMetrics {
    /** The height the line box actually got. */
    readonly lineHeightPx: number;
    /** What single spacing would have been for the tallest run on the line. */
    readonly naturalHeightPx: number;
    /** How far the deepest descender reaches below the baseline, as a POSITIVE length. */
    readonly descentPx: number;
    readonly rule: LineRule;
}

/**
 * How far below the top of its line box the baseline sits.
 *
 * ## Where the extra leading goes is what separates the rules
 *
 * `atLeast` puts it ABOVE: the baseline stays one descender off the bottom, so
 * text sinks to the foot of a taller line. `auto` puts it BELOW: the baseline
 * stays exactly where single spacing would have put it and the line grows
 * downwards, which is why 1.5-spaced text does not drift down the page. `exact`
 * ignores the font altogether.
 */
export function baselineOffsetPx(metrics: BaselineMetrics): number {
    if ('exact' === metrics.rule) {
        return EXACT_BASELINE_RATIO * metrics.lineHeightPx;
    }

    // `w:docGrid` puts the leading round the line rather than on one side of
    // it: measured against LibreOffice, an 18pt pitch moved a 11.5pt line's
    // first baseline from 804.14 to 800.89 — 3.25 down, which is half of the
    // 6.5 the grid added. The text sits in the middle of its grid line.
    //
    // There is no leading to halve when the line is TALLER than the pitch, and
    // half of a negative number puts the baseline above the box. Measured: a
    // line carrying a 36pt picture on an 18pt grid printed its baseline 36.00
    // below the top of a box three grid lines deep — the line flush with the
    // top and all the spare room below it.
    if ('grid' === metrics.rule) {
        return Math.max(0, (metrics.lineHeightPx - metrics.naturalHeightPx) / 2)
            + metrics.naturalHeightPx - metrics.descentPx;
    }

    // `atLeast` measures from the bottom of the box it was given; `auto`
    // measures from the bottom of the box the font asked for.
    const from = 'atLeast' === metrics.rule ? metrics.lineHeightPx : metrics.naturalHeightPx;

    return from - metrics.descentPx;
}
