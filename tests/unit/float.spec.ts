import { describe, expect, it } from 'vitest';

import { lineBoxAt, type LineBox, type PlacedFloat } from '../../src/layout/float.js';

/** A column 400 wide starting at 100, so neither edge is zero. */
const COLUMN: LineBox = { leftPx: 100, widthPx: 400 };

const float = (
    xPx: number,
    widthPx: number,
    yPx = 0,
    heightPx = 50,
    wrap: 'square' | 'none' = 'square',
): PlacedFloat => ({
    image: { content: { bytes: new Uint8Array(), contentType: 'image/png' }, widthPx, heightPx },
    xPx,
    yPx,
    widthPx,
    heightPx,
    wrap,
    behindText: false,
    exclusion: { xPx, yPx, widthPx, heightPx },
});

describe('lineBoxAt', () => {
    it('leaves a line the whole column when nothing is in the way', () => {
        expect(lineBoxAt(COLUMN, 0, 20, [])).toEqual(COLUMN);
    });

    it('puts the line to the RIGHT of a float near the left edge', () => {
        // 150 of room on the left against 300 on the right: the wider side
        // wins, and the line starts where the float ends.
        expect(lineBoxAt(COLUMN, 0, 20, [float(150, 50)]))
            .toEqual({ leftPx: 200, widthPx: 300 });
    });

    it('keeps the line LEFT of a float near the right edge', () => {
        expect(lineBoxAt(COLUMN, 0, 20, [float(350, 50)]))
            .toEqual({ leftPx: 100, widthPx: 250 });
    });

    it('gives a line no room at all when a float spans the column', () => {
        // Not a negative width, which would break the line after every
        // character rather than moving it past the float.
        expect(lineBoxAt(COLUMN, 0, 20, [float(50, 500)]))
            .toEqual({ leftPx: 100, widthPx: 0 });
    });

    it('ignores a float that does not reach this line', () => {
        // The float ends at 50; a line starting there is clear of it. Half-open
        // at both ends, which is the line LibreOffice gives the column back to.
        expect(lineBoxAt(COLUMN, 50, 20, [float(150, 50, 0, 50)])).toEqual(COLUMN);
        expect(lineBoxAt(COLUMN, 49, 20, [float(150, 50, 0, 50)]))
            .toEqual({ leftPx: 200, widthPx: 300 });
    });

    it('ignores a line that ends before the float begins', () => {
        expect(lineBoxAt(COLUMN, 0, 20, [float(150, 50, 20, 50)])).toEqual(COLUMN);
    });

    it('steps aside for nothing when the float wraps NONE', () => {
        // It sits over or under the text; the text does not move.
        expect(lineBoxAt(COLUMN, 0, 20, [float(150, 50, 0, 50, 'none')])).toEqual(COLUMN);
    });

    it('takes every float in turn', () => {
        // One at each edge: what is left is the band between them.
        const box = lineBoxAt(COLUMN, 0, 20, [float(100, 60), float(440, 60)]);

        expect(box).toEqual({ leftPx: 160, widthPx: 280 });
    });

    it('keeps the clear space around a float, not just the picture', () => {
        // The exclusion is what text avoids; the picture is only what is drawn.
        const spaced: PlacedFloat = { ...float(150, 50), exclusion: { xPx: 140, yPx: 0, widthPx: 70, heightPx: 50 } };

        expect(lineBoxAt(COLUMN, 0, 20, [spaced])).toEqual({ leftPx: 210, widthPx: 290 });
    });
});
