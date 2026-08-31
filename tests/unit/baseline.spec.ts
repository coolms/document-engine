import { describe, expect, it } from 'vitest';

import { baselineOffsetPx, EXACT_BASELINE_RATIO } from '../../src/layout/baseline.js';

/**
 * A font 20 tall naturally, reaching 4 below the baseline — so its natural
 * baseline sits at 16, and every case below moves visibly off that.
 */
const FONT = { naturalHeightPx: 20, descentPx: 4 };

describe('baselineOffsetPx', () => {
    describe('auto — the extra leading goes BELOW', () => {
        it('leaves the baseline where single spacing put it', () => {
            // 1.5 spacing makes the box 30 tall and does NOT move the text down;
            // if it did, every 1.5-spaced document would drift a line per page.
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 30, rule: 'auto' })).toBe(16);
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 40, rule: 'auto' })).toBe(16);
        });

        it('is the natural baseline when nothing overrode the height', () => {
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 20, rule: 'auto' })).toBe(16);
        });
    });

    describe('atLeast — the extra leading goes ABOVE', () => {
        it('keeps the baseline one descender off the BOTTOM', () => {
            // Text sinks to the foot of a taller line, which is the visible
            // difference from `auto` at the same height.
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 30, rule: 'atLeast' })).toBe(26);
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 40, rule: 'atLeast' })).toBe(36);
        });

        it('agrees with auto when the line was not enlarged', () => {
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 20, rule: 'atLeast' })).toBe(16);
        });
    });

    describe('exact — a proportion, not a metric', () => {
        it('takes four fifths of the line whatever the font', () => {
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 30, rule: 'exact' })).toBe(24);
            // A font twice as deep with a different natural height lands in the
            // same place: verified against LibreOffice across three faces.
            expect(baselineOffsetPx({
                naturalHeightPx: 11, descentPx: 8, lineHeightPx: 30, rule: 'exact',
            })).toBe(24);
        });

        it('uses the ratio it publishes', () => {
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 50, rule: 'exact' }))
                .toBe(EXACT_BASELINE_RATIO * 50);
        });

        it('can put the baseline ABOVE where the font wanted it', () => {
            // An exact line shorter than the font asks for loses the descender
            // to the line below rather than moving the baseline — that is what
            // "exact" means, and why it can clip.
            expect(baselineOffsetPx({ ...FONT, lineHeightPx: 10, rule: 'exact' })).toBe(8);
        });
    });

    it('separates the three rules at one and the same height', () => {
        // The whole reason the rule is carried alongside the number: these are
        // three different answers for one identical line box.
        const at = (rule: 'auto' | 'atLeast' | 'exact'): number =>
            baselineOffsetPx({ ...FONT, lineHeightPx: 30, rule });

        expect([at('auto'), at('atLeast'), at('exact')]).toEqual([16, 26, 24]);
    });
});
