import { describe, expect, it } from 'vitest';

import { alignLine } from '../../src/layout/alignment.js';
import type { Line } from '../../src/layout/line-breaker.js';

/**
 * A line of a given width, with a given number of stretchable gaps.
 *
 * Built by hand rather than broken from text so each case states the ONE thing
 * it is about: the alignment arithmetic, not the measuring that produced the
 * width.
 */
function line(widthPx: number, spaceGaps = 0, endedByMandatoryBreak = false): Line {
    return {
        pieces: [],
        widthPx,
        overflows: false,
        endedByMandatoryBreak,
        startsAt: 0,
        endsAt: 0,
        spaceGaps,
    };
}

describe('alignLine', () => {
    describe('where the line sits', () => {
        it('leaves a left-aligned line at the start of its column', () => {
            expect(alignLine(line(40), 100, 'left')).toEqual({ offsetPx: 0, wordSpacingPx: 0 });
        });

        it('treats an unstated alignment as left', () => {
            // A paragraph with no `w:jc` is left-aligned, and the default has to
            // live in one place or a caller that omits it gets something else.
            expect(alignLine(line(40), 100)).toEqual({ offsetPx: 0, wordSpacingPx: 0 });
        });

        it('centres on HALF the slack', () => {
            expect(alignLine(line(40), 100, 'center').offsetPx).toBe(30);
        });

        it('pushes a right-aligned line by ALL of it', () => {
            expect(alignLine(line(40), 100, 'right').offsetPx).toBe(60);
        });

        it('puts a line that exactly fills its column at the start, whatever the alignment', () => {
            for (const alignment of ['left', 'center', 'right', 'justify'] as const) {
                expect(alignLine(line(100), 100, alignment).offsetPx).toBe(0);
            }
        });

        it('never moves an OVERFLOWING line further out', () => {
            // A long URL is wider than its column, so the slack is negative.
            // Centring on that would put the text LEFT of the indent — Word lets
            // it overflow to the right instead.
            expect(alignLine(line(140), 100, 'center').offsetPx).toBe(0);
            expect(alignLine(line(140), 100, 'right').offsetPx).toBe(0);
        });
    });

    describe('justification', () => {
        it('shares the slack out over the gaps', () => {
            expect(alignLine(line(40, 3), 100, 'justify'))
                .toEqual({ offsetPx: 0, wordSpacingPx: 20 });
        });

        it('leaves the LAST line of the paragraph alone', () => {
            // The most recognisable way justified text goes wrong: a final line
            // of three words stretched across the whole column.
            expect(alignLine(line(40, 3), 100, 'justify', true).wordSpacingPx).toBe(0);
        });

        it('leaves a line the AUTHOR ended alone too', () => {
            // A `w:br` ends the line for the same reason a paragraph end does —
            // there is no more text coming to fill it.
            expect(alignLine(line(40, 3, true), 100, 'justify').wordSpacingPx).toBe(0);
        });

        it('has nothing to stretch on a single-word line', () => {
            // No gaps: the slack has nowhere to go, and dividing by zero would
            // put Infinity on the line.
            expect(alignLine(line(40, 0), 100, 'justify').wordSpacingPx).toBe(0);
        });

        it('does not SHRINK an overflowing line', () => {
            expect(alignLine(line(140, 3), 100, 'justify').wordSpacingPx).toBe(0);
        });

        it('does not move the start of the line', () => {
            expect(alignLine(line(40, 3), 100, 'justify').offsetPx).toBe(0);
        });
    });
});
