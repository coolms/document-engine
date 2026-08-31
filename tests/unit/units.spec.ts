import { describe, expect, it } from 'vitest';

import {
    EMU_PER_INCH,
    POINTS_PER_INCH,
    PX_PER_INCH,
    TWIPS_PER_INCH,
    eighthPointsToPoints,
    emuToPoints,
    emuToPx,
    halfPointsToPoints,
    halfPointsToPx,
    pointsToEmu,
    pointsToHalfPoints,
    pointsToPx,
    pointsToTwips,
    pxToEmu,
    pxToHalfPoints,
    pxToPoints,
    pxToTwips,
    twipsToPoints,
    twipsToPx,
} from '../../src/ooxml/units.js';

describe('units', () => {
    describe('against the definitions, not against the code', () => {
        // Each expectation is stated as the DEFINITION of the unit, so the test
        // disagrees with the implementation rather than restating it.
        it('measures an inch the same way in every unit', () => {
            expect(twipsToPoints(TWIPS_PER_INCH)).toBe(POINTS_PER_INCH);
            expect(twipsToPx(TWIPS_PER_INCH)).toBe(PX_PER_INCH);
            expect(emuToPoints(EMU_PER_INCH)).toBe(POINTS_PER_INCH);
            expect(emuToPx(EMU_PER_INCH)).toBe(PX_PER_INCH);
            expect(pointsToPx(POINTS_PER_INCH)).toBe(PX_PER_INCH);
            expect(pxToPoints(PX_PER_INCH)).toBe(POINTS_PER_INCH);
        });

        it('reads US Letter out of a real w:pgSz', () => {
            // 12240 x 15840 twips is what the Word-authored fixture declares.
            expect(twipsToPoints(12240)).toBe(612); // 8.5in
            expect(twipsToPoints(15840)).toBe(792); // 11in
            expect(twipsToPx(1440)).toBe(96); // a one-inch margin
        });

        it('knows w:sz means HALF-points on a run and EIGHTH-points on a border', () => {
        // Defined by the OOXML schema, not measured — the two units share
        // an attribute name and nothing else.
            // The same attribute name, two units. The Word fixture's heading
            // carries w:sz="28", which is 14pt and not 28pt or 3.5pt.
            expect(halfPointsToPoints(28)).toBe(14);
            expect(eighthPointsToPoints(28)).toBe(3.5);
            expect(halfPointsToPx(24)).toBe(16); // 12pt body text
        });
    });

    describe('float behaviour', () => {
        it('round-trips every twip value in a page dimension exactly', () => {
            for (let twips = 0; twips <= 20000; twips++) {
                expect(pointsToTwips(twipsToPoints(twips))).toBe(twips);
            }
        });

        it('converts twips by DIVIDING, which the rounding above would hide', () => {
            // `twips * 0.05` and `twips / 20` disagree for 7186 of these values
            // because 0.05 has no exact binary form. The round trip above
            // cannot see it — pointsToTwips rounds, and rounding absorbs the
            // error — so this states the property WITHOUT any rounding.
            //
            // It matters because layout compares rather than rounds: line
            // breaking asks `width <= available`, and an epsilon on the wrong
            // side of that ends a line one word early.
            for (let twips = 1; twips <= 20000; twips++) {
                expect(twipsToPoints(twips) * 20).toBe(twips);
            }
        });

        it('round-trips through PIXELS too, which needs the rounding', () => {
            // twips -> px -> twips carries about 1.5e-11 of error. Without the
            // rounding this emits w:w="12240.000000000002", and Word rejects a
            // non-integer measurement.
            for (const twips of [1, 567, 1440, 12240, 15840, 19999]) {
                expect(pxToTwips(twipsToPx(twips))).toBe(twips);
            }
        });

        it('returns INTEGERS for every unit that is written back to the file', () => {
            // OOXML attribute values are integers. A fractional one produces a
            // document Word refuses to open, which is a failure at the user's
            // desk rather than here.
            expect(Number.isInteger(pointsToTwips(11.3))).toBe(true);
            expect(Number.isInteger(pxToTwips(37.7))).toBe(true);
            expect(Number.isInteger(pointsToHalfPoints(10.7))).toBe(true);
            expect(Number.isInteger(pxToHalfPoints(13.1))).toBe(true);
            expect(Number.isInteger(pointsToEmu(9.37))).toBe(true);
            expect(Number.isInteger(pxToEmu(15.9))).toBe(true);
        });

        it('converts points and pixels without any loss at all', () => {
            // 1pt = 96/72 = 4/3 px, and both are exactly representable, so this
            // direction needs no rounding and must not silently acquire any.
            for (let points = 0; points <= 2000; points++) {
                expect(pxToPoints(pointsToPx(points))).toBe(points);
            }
        });

        it('rounds half-points to the nearest, not towards zero', () => {
            expect(pointsToHalfPoints(10.3)).toBe(21);
            expect(pointsToTwips(1.04)).toBe(21);
        });
    });

    describe('direction and sign', () => {
        it('keeps negative measurements negative', () => {
            // A hanging indent is a negative w:firstLine, and an absolute value
            // would turn it into an ordinary indent pointing the wrong way.
            expect(twipsToPx(-720)).toBe(-48);
            expect(pxToTwips(-48)).toBe(-720);
            expect(emuToPx(-914400)).toBe(-96);
        });

        it('leaves zero alone in both directions', () => {
            expect(twipsToPx(0)).toBe(0);
            expect(pxToTwips(0)).toBe(0);
            expect(halfPointsToPx(0)).toBe(0);
        });

        it('scales linearly', () => {
            expect(twipsToPx(2880)).toBe(twipsToPx(1440) * 2);
            expect(emuToPoints(EMU_PER_INCH * 3)).toBe(emuToPoints(EMU_PER_INCH) * 3);
        });
    });
});
