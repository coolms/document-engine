import { describe, expect, it } from 'vitest';

import { formatNumeral, letters, roman } from '../../src/text/numerals.js';

describe('roman', () => {
    it('writes the numerals a document actually uses', () => {
        expect([1, 2, 3, 4, 5, 9, 10].map((n) => roman(n, false)))
            .toEqual(['i', 'ii', 'iii', 'iv', 'v', 'ix', 'x']);
    });

    it('subtracts rather than repeating four times', () => {
        // iv, not iiii — and the same at every power.
        expect([4, 9, 40, 90, 400, 900].map((n) => roman(n, false)))
            .toEqual(['iv', 'ix', 'xl', 'xc', 'cd', 'cm']);
    });

    it('upper-cases the whole numeral, not its first letter', () => {
        expect(roman(38, true)).toBe('XXXVIII');
    });

    it('gives nothing for a number below one', () => {
        // A page count of zero has no numeral, and inventing one would print a
        // numeral where the document has no page.
        expect(roman(0, false)).toBe('');
    });
});

describe('letters', () => {
    it('runs a to z', () => {
        expect([1, 2, 26].map((n) => letters(n, false))).toEqual(['a', 'b', 'z']);
    });

    it('REPEATS the letter past z, as Word does', () => {
        // Word's sequence is aa, bb, cc — not the spreadsheet's aa, ab, ac.
        expect([27, 28, 53].map((n) => letters(n, false))).toEqual(['aa', 'bb', 'aaa']);
    });

    it('upper-cases every letter of a repeat', () => {
        expect(letters(27, true)).toBe('AA');
    });

    it('gives nothing for a number below one', () => {
        expect(letters(0, false)).toBe('');
    });
});

describe('formatNumeral', () => {
    it('writes the same number every way a document can ask for', () => {
        // Three is the smallest number that tells all five apart.
        expect([
            formatNumeral(3, 'decimal'),
            formatNumeral(3, 'lowerRoman'),
            formatNumeral(3, 'upperRoman'),
            formatNumeral(3, 'lowerLetter'),
            formatNumeral(3, 'upperLetter'),
        ]).toEqual(['3', 'iii', 'III', 'c', 'C']);
    });
});
