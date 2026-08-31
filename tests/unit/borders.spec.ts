import { describe, expect, it } from 'vitest';

import {
    resolveCellBorders,
    strongerBorder,
    type BorderSide,
    type BoxBorders,
} from '../../src/layout/borders.js';
import { withSide } from '../../src/layout/table-layout.js';

const side = (widthPx: number): BorderSide => ({ widthPx, colorHex: '#000000', style: 'solid' });

/** Every side a different width, so no two can be confused for one another. */
const TABLE: BoxBorders = {
    top: side(1),
    left: side(2),
    bottom: side(3),
    right: side(4),
    insideH: side(5),
    insideV: side(6),
};

const MIDDLE = { firstRow: false, lastRow: false, firstColumn: false, lastColumn: false };
const ALONE = { firstRow: true, lastRow: true, firstColumn: true, lastColumn: true };

const widths = (borders: BoxBorders): (number | undefined)[] =>
    [borders.top?.widthPx, borders.bottom?.widthPx, borders.left?.widthPx, borders.right?.widthPx];

describe('resolveCellBorders', () => {
    it('gives a cell in the middle the INSIDE borders on every side', () => {
        expect(widths(resolveCellBorders(TABLE, undefined, MIDDLE))).toEqual([5, 5, 6, 6]);
    });

    it('gives the only cell in a table its OUTER borders on every side', () => {
        expect(widths(resolveCellBorders(TABLE, undefined, ALONE))).toEqual([1, 3, 2, 4]);
    });

    it('takes each side from the edge that side is on', () => {
        // A cell on the first row and last column: outer above and to the
        // right, inside below and to the left.
        const at = { firstRow: true, lastRow: false, firstColumn: false, lastColumn: true };

        expect(widths(resolveCellBorders(TABLE, undefined, at))).toEqual([1, 5, 6, 4]);
    });

    it('lets the CELL beat the table on the side it declares, and only there', () => {
        const cell: BoxBorders = { left: side(9) };

        expect(widths(resolveCellBorders(TABLE, cell, MIDDLE))).toEqual([5, 5, 9, 6]);
    });

    it('lets a cell declare a border where the table has none', () => {
        expect(widths(resolveCellBorders(undefined, { top: side(7) }, MIDDLE)))
            .toEqual([7, undefined, undefined, undefined]);
    });

    it('leaves a side absent when neither declares it', () => {
        const sparse: BoxBorders = { insideH: side(5) };
        const resolved = resolveCellBorders(sparse, undefined, MIDDLE);

        expect(resolved.left).toBeUndefined();
        expect(resolved.right).toBeUndefined();
        // And an absent side is absent, not present-and-zero: a renderer tests
        // for the key.
        expect('left' in resolved).toBe(false);
    });

    it('never hands a cell the inside borders as its own', () => {
        // `insideH`/`insideV` describe the table's interior; a resolved cell has
        // four sides and no interior at all.
        const resolved = resolveCellBorders(TABLE, undefined, MIDDLE);

        expect(resolved.insideH).toBeUndefined();
        expect(resolved.insideV).toBeUndefined();
    });
});

describe('strongerBorder', () => {
    it('gives the edge to the HEAVIER of the two', () => {
        expect(strongerBorder(side(4), side(1))?.widthPx).toBe(4);
        expect(strongerBorder(side(1), side(4))?.widthPx).toBe(4);
    });

    it('gives a TIE to the later one', () => {
        // The right-hand cell's left border, the lower row's top border.
        // Measured: two 2pt rules of different colours resolve to the later
        // cell's, and the same pair with the colours swapped resolves the other
        // way — which is what rules out a rule about colour.
        const earlier: BorderSide = { widthPx: 2, colorHex: '#FF0000', style: 'solid' };
        const later: BorderSide = { widthPx: 2, colorHex: '#0000FF', style: 'solid' };

        expect(strongerBorder(earlier, later)).toBe(later);
        expect(strongerBorder(later, earlier)).toBe(earlier);
    });

    it('lets a declared border beat one that is absent', () => {
        expect(strongerBorder(undefined, side(2))?.widthPx).toBe(2);
        expect(strongerBorder(side(2), undefined)?.widthPx).toBe(2);
    });

    it('leaves an edge neither side asked for alone', () => {
        expect(strongerBorder(undefined, undefined)).toBeUndefined();
    });
});

describe('withSide', () => {
    it('replaces one side and leaves the others alone', () => {
        const result = withSide({ top: side(1), left: side(2) }, 'top', side(9));

        expect(result.top?.widthPx).toBe(9);
        expect(result.left?.widthPx).toBe(2);
    });

    it('REMOVES a side rather than leaving the key holding undefined', () => {
        // The whole model is "absent means no border". A key present and
        // undefined draws the same today, and would stop doing so the first
        // time anything asked whether the side was declared.
        const result = withSide({ top: side(1), left: side(2) }, 'top', undefined);

        expect('top' in result).toBe(false);
        expect('left' in result).toBe(true);
    });
});
