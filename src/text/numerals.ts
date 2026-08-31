/**
 * Numbers written as something other than digits.
 *
 * Shared by list numbering and by field switches, because they are the same
 * question asked twice: a list level of `lowerRoman` and a `PAGE \* roman`
 * field both want the third thing called `iii`. Two implementations would
 * eventually disagree, and the disagreement would show as a list numbered
 * differently from the page it is on.
 */

export type NumeralStyle =
    | 'decimal'
    | 'lowerLetter'
    | 'upperLetter'
    | 'lowerRoman'
    | 'upperRoman';

export function formatNumeral(value: number, style: NumeralStyle): string {
    switch (style) {
        case 'lowerLetter':
            return letters(value, false);
        case 'upperLetter':
            return letters(value, true);
        case 'lowerRoman':
            return roman(value, false);
        case 'upperRoman':
            return roman(value, true);
        default:
            return String(value);
    }
}

/**
 * Word's letter sequence: a…z, then aa, bb, cc — the letter REPEATED, not the
 * spreadsheet's aa, ab, ac.
 */
export function letters(value: number, upper: boolean): string {
    if (value < 1) {
        return '';
    }

    const index = (value - 1) % 26;
    const repeat = Math.floor((value - 1) / 26) + 1;
    const letter = String.fromCharCode('a'.charCodeAt(0) + index);

    return (upper ? letter.toUpperCase() : letter).repeat(repeat);
}

const ROMAN: readonly (readonly [number, string])[] = [
    [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'],
    [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'],
    [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
];

export function roman(value: number, upper: boolean): string {
    let remaining = value;
    let out = '';
    for (const [amount, numeral] of ROMAN) {
        while (remaining >= amount) {
            out += numeral;
            remaining -= amount;
        }
    }

    return upper ? out.toUpperCase() : out;
}
