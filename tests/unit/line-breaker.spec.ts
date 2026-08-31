import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import type { BorderSide, BorderStyle } from '../../src/layout/borders.js';
import { breakIntoLines, type Line, type StyledRun } from '../../src/layout/line-breaker.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');

function font(file: string): TrueTypeFont {
    return TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, file))));
}

const MONO = font('LiberationMono-Regular.ttf');
const SERIF = font('LiberationSerif-Regular.ttf');

/** One monospaced character at 16px — the unit every width below is built from. */
const CELL = MONO.measureAdvance('x', 16).widthPt;

function mono(text: string): StyledRun[] {
    return [{ text, font: MONO, sizePx: 16 }];
}

function textOf(line: Line): string {
    return line.pieces.map((piece) => piece.text).join('');
}

/**
 * A monospaced font makes line breaking arithmetic exact: a column of N cells
 * fits N characters and not one more. That turns "did it break in the right
 * place" from a judgement into a calculation, which is the only way to assert
 * it without pinning recorded pixel values that no one can check.
 */
describe('breakIntoLines', () => {
    describe('fitting', () => {
        it('fills a line to exactly the column width before breaking', () => {
            // "aaa bbb" is 7 cells. At 7 cells both words fit; at 6 they cannot.
            expect(breakIntoLines(mono('aaa bbb'), CELL * 7).map(textOf)).toEqual(['aaa bbb']);
            expect(breakIntoLines(mono('aaa bbb'), CELL * 6).map(textOf)).toEqual(['aaa ', 'bbb']);
        });

        it('does not let a trailing space push the break one word early', () => {
            // The space after "aaa" ends the line and hangs into the margin. If
            // it counted toward the fit, "bbb" would be pushed to line 3 in a
            // column that plainly has room for it.
            const lines = breakIntoLines(mono('aaa bbb'), CELL * 7);

            expect(lines).toHaveLength(1);
            expect(lines[0]!.widthPx).toBeCloseTo(CELL * 7, 6);
        });

        it('reports the fitting width without trailing spaces', () => {
            const [line] = breakIntoLines(mono('aa   '), CELL * 10);

            // Five characters of ink and space, but only two count as width.
            expect(line!.widthPx).toBeCloseTo(CELL * 2, 6);
        });

        it('keeps a wrapped line from starting with the previous line\'s space', () => {
            const lines = breakIntoLines(mono('aaa bbb'), CELL * 4);

            expect(lines.map(textOf)).toEqual(['aaa ', 'bbb']);
            expect(textOf(lines[1]!).startsWith(' ')).toBe(false);
        });
    });

    describe('unbreakable text', () => {
        it('CHOPS an over-long word at the last character that fits', () => {
            // This engine used to place it whole and flag the overflow, on the
            // reasoning that a mid-word break is one the author never typed.
            // Measured, that is what Word and LibreOffice both do
            // rather than let text run off the page: `gggggggghhhhhhhh` printed
            // as `ggggggggh` and `hhhhhhh` in a 51.3pt measure.
            //
            // Five monospaced cells, so five characters a line and no doubt
            // about where the cut falls.
            const lines = breakIntoLines(mono('supercalifragilistic'), CELL * 5);

            expect(lines.map(textOf)).toEqual(['super', 'calif', 'ragil', 'istic']);
            expect(lines.every((line) => !line.overflows)).toBe(true);
        });

        it('never cuts INTO a picture, which has one character and no glyphs', () => {
            // A picture's piece stands for itself with a single
            // object-replacement character, and its width comes from the
            // picture rather than from measuring that character. Cutting into
            // it would measure the character — a couple of points where the
            // picture is a hundred — and draw the picture at that width.
            //
            // No printed page decides this one: it is what the piece MEANS.
            const image = {
                content: { bytes: new Uint8Array(), contentType: 'image/png' },
                widthPx: CELL * 20,
                heightPx: CELL * 2,
            };
            const lines = breakIntoLines(
                [{ text: '￼', font: MONO, sizePx: 16, image }],
                CELL * 5,
            );

            expect(lines).toHaveLength(1);
            expect(lines[0]!.pieces[0]!.widthPx).toBe(CELL * 20);
            expect(lines[0]!.overflows).toBe(true);
        });

        it('leaves the overflow flag for what cannot be chopped at all', () => {
            // One character wider than its own line: there is nothing to cut
            // into, and cutting nothing off would put the same word back for
            // ever. That is the case the flag still exists for.
            const lines = breakIntoLines(mono('abc'), CELL / 2);

            expect(lines[0]!.overflows).toBe(true);
        });

        it('still breaks the rest of the paragraph around it', () => {
            const lines = breakIntoLines(mono('hi supercalifragilistic ok'), CELL * 5);

            // The last chopped line keeps the space that ended it, as every
            // wrapped line here does.
            expect(lines.map(textOf)).toEqual(['hi ', 'super', 'calif', 'ragil', 'istic ', 'ok']);
        });
    });

    describe('author breaks', () => {
        it('ends the line where the author typed a newline', () => {
            const lines = breakIntoLines(mono('a\nb'), CELL * 100);

            expect(lines.map(textOf)).toEqual(['a', 'b']);
            expect(lines[0]!.endedByMandatoryBreak).toBe(true);
            expect(lines[1]!.endedByMandatoryBreak).toBe(false);
        });

        it('treats CRLF as ONE break, not two', () => {
            // Two lines, not three with an empty one between — the classic
            // Windows-file symptom.
            expect(breakIntoLines(mono('a\r\nb'), CELL * 100).map(textOf)).toEqual(['a', 'b']);
        });

        it('preserves a deliberately empty line', () => {
            expect(breakIntoLines(mono('a\n\nb'), CELL * 100).map(textOf)).toEqual(['a', '', 'b']);
        });

        it('never measures the break character itself', () => {
            const [line] = breakIntoLines(mono('ab\n'), CELL * 100);

            expect(line!.widthPx).toBeCloseTo(CELL * 2, 6);
        });
    });

    describe('break opportunities', () => {
        it('breaks after a hyphen but not inside the word', () => {
            expect(breakIntoLines(mono('well-known'), CELL * 6).map(textOf)).toEqual(['well-', 'known']);
        });

        it('refuses to break at a NO-BREAK SPACE, where an ordinary space breaks', () => {
            // The author wrote U+00A0 to say "not here"; a breaker treating it
            // as an ordinary space overrides them silently.
            //
            // The CONTROL is the point. Asserting only that the NBSP case
            // yields one line proves nothing by itself — it would pass just as
            // well if the column were wide enough, or if nothing ever broke.
            // Mutation testing caught exactly that. The identical string with
            // U+0020 must break in the same column, so the difference can only
            // be the space.
            const nonBreaking = breakIntoLines(mono('10 kg'), CELL * 3);
            const ordinary = breakIntoLines(mono('10 kg'), CELL * 3);

            //
            // A word too long for its line is CHOPPED rather than left whole,
            // so what the no-break space protects is no longer the string but
            // the JOIN: the chop goes round it, and the space is never the
            // last thing on a line nor the first thing on the next.
            const cuts = nonBreaking.map(textOf);

            expect(cuts.join('')).toBe('10 kg');
            expect(cuts.some((line) => line.endsWith('0'))).toBe(false);
            expect(cuts.some((line) => line.startsWith(' '))).toBe(false);
            // And not at the far side of it either: a line ENDING with the
            // space is the same forbidden join, cut from the other end. Only
            // asserting the two above let a chop that ignored the rule
            // entirely pass, because it put the space at a line's end.
            expect(cuts.some((line) => line.endsWith(' '))).toBe(false);

            expect(ordinary.map(textOf)).toEqual(['10 ', 'kg']);
            expect(ordinary[0]!.overflows).toBe(false);
        });

        /**
         * A monospaced font makes this exact rather than approximate: every
         * glyph is one cell WIDE, `.notdef` included. So a character the font
         * has no glyph for used to cost exactly one cell — and the assertion
         * below is "the same string still fits the same column", with the
         * control proving the column is the tight one.
         *
         * Liberation Mono has no glyph for U+200B, U+2060 or U+FEFF, which is
         * why they are the ones tested here.
         */
        it('does not let an INVISIBLE character push the break a word early', () => {
            const ZWSP = String.fromCodePoint(0x200b);
            const WJ = String.fromCodePoint(0x2060);
            const BOM = String.fromCodePoint(0xfeff);

            // The control: 7 cells of text in a 7-cell column is the tight fit
            // that one phantom cell would break. Asserting only the cases below
            // would pass just as well in a column with room to spare.
            expect(breakIntoLines(mono('aaa bbb'), CELL * 7).map(textOf)).toEqual(['aaa bbb']);
            expect(breakIntoLines(mono('aaa bbb'), CELL * 6).map(textOf)).toEqual(['aaa ', 'bbb']);

            for (const invisible of [ZWSP, WJ, BOM]) {
                // Placed inside the first word, where it can only add width —
                // it offers no break there that would rescue the line.
                const lines = breakIntoLines(mono('a' + invisible + 'aa bbb'), CELL * 7);

                expect(lines.map(textOf).join('')).toBe('a' + invisible + 'aa bbb');
                expect(lines.length, 'an invisible character took a cell').toBe(1);
            }
        });

        it('still lets a ZERO WIDTH SPACE offer the break it exists for', () => {
            // Costing nothing must not mean doing nothing: U+200B is a break
            // opportunity, and the proof is that it CHANGES where the line
            // ends. Without it an over-long word is chopped where the width
            // runs out; with it the break lands where the author put it.
            const ZWSP = String.fromCodePoint(0x200b);

            expect(breakIntoLines(mono('aaaabbbb'), CELL * 6).map(textOf))
                .toEqual(['aaaabb', 'bb']);
            expect(breakIntoLines(mono('aaaa' + ZWSP + 'bbbb'), CELL * 6).map(textOf))
                .toEqual(['aaaa' + ZWSP, 'bbbb']);
        });

        it('does not charge LETTER SPACING for a character that is not a letter', () => {
            // Tracking is added per character, so an invisible one puts the
            // phantom width straight back for exactly the runs that set
            // `w:spacing` — the font skipping it is not enough on its own.
            const ZWSP = String.fromCodePoint(0x200b);
            const tracked = (text: string): StyledRun[] =>
                [{ text, font: MONO, sizePx: 16, letterSpacingPx: 4 }];

            expect(breakIntoLines(tracked('aaaa' + ZWSP + 'bbbb'), CELL * 100)[0]!.widthPx)
                .toBe(breakIntoLines(tracked('aaaabbbb'), CELL * 100)[0]!.widthPx);
        });

        it('breaks between ideographs, which have no spaces to break at', () => {
            // Without this a CJK paragraph is one unbreakable word and overflows
            // the page entirely.
            const lines = breakIntoLines([{ text: '日本語文書', font: SERIF, sizePx: 16 }], SERIF.measureAdvance('日本', 16).widthPt);

            expect(lines.length).toBeGreaterThan(1);
        });
    });

    describe('mixed styles', () => {
        it('measures each piece with its own run\'s font and size', () => {
            // The same letters at two sizes are not the same width; a breaker
            // that measured the paragraph with one font would break in a place
            // that looks arbitrary in the rendered document.
            const runs: StyledRun[] = [
                { text: 'aa ', font: MONO, sizePx: 16 },
                { text: 'bb', font: MONO, sizePx: 32 },
            ];

            const [line] = breakIntoLines(runs, CELL * 100);

            expect(line!.pieces).toHaveLength(2);
            expect(line!.pieces[1]!.widthPx).toBeCloseTo(line!.pieces[0]!.widthPx / 3 * 2 * 2, 6);
        });

        it('splits a word that spans a style change into a piece per run', () => {
            const runs: StyledRun[] = [
                { text: 'wo', font: MONO, sizePx: 16 },
                { text: 'rd', font: SERIF, sizePx: 16 },
            ];

            const [line] = breakIntoLines(runs, CELL * 100);

            expect(line!.pieces.map((p) => p.text)).toEqual(['wo', 'rd']);
            expect(line!.pieces.map((p) => p.runIndex)).toEqual([0, 1]);
        });
    });

    describe('degenerate input', () => {
        it('returns one empty line for empty text rather than nothing', () => {
            // A paragraph with no text still occupies a line's height on the
            // page; returning zero lines would collapse it.
            const lines = breakIntoLines(mono(''), CELL * 10);

            expect(lines).toHaveLength(1);
            expect(lines[0]!.widthPx).toBe(0);
        });
    });

    describe('source offsets', () => {
        /**
         * Every line reports the span of source it covers. An editor maps a page
         * boundary back to a caret position with this, so an off-by-one puts the
         * caret in the wrong word — and the pieces alone cannot supply it,
         * because a mandatory break occupies source and appears in no piece.
         */
        const withoutBreaks = (text: string): string => text.replace(/[\r\n]/g, '');

        it('spans exactly the characters the line drew', () => {
            // Self-checking: slice the source with the reported span and it must
            // be the line's own text, once the breaks that are positions rather
            // than ink are taken out.
            const text = 'the quick brown fox jumps over the lazy dog';
            for (const width of [5, 7, 10, 13, 21, 40]) {
                for (const line of breakIntoLines(mono(text), CELL * width)) {
                    expect(withoutBreaks(text.slice(line.startsAt, line.endsAt))).toBe(textOf(line));
                }
            }
        });

        it('covers the whole text with no gap and no overlap', () => {
            // A gap loses characters and an overlap double-counts them; either
            // makes a caret map that is wrong somewhere in the middle and looks
            // right at both ends.
            // Both line-ending conventions, because a CRLF is one break and TWO
            // units — the case where a gap is easiest to introduce and hardest
            // to notice, since the ends of the document still look right.
            for (const text of [
                'alpha beta gamma delta epsilon zeta eta theta',
                'alpha beta\ngamma delta\nepsilon zeta',
                'alpha beta\r\ngamma delta\r\nepsilon zeta',
                'alpha\n\r\nbeta',
            ]) {
                const lines = breakIntoLines(mono(text), CELL * 11);

                expect(lines[0]!.startsAt).toBe(0);
                expect(lines[lines.length - 1]!.endsAt).toBe(text.length);
                for (let index = 1; index < lines.length; index++) {
                    expect(lines[index]!.startsAt).toBe(lines[index - 1]!.endsAt);
                }
            }
        });

        it('counts the mandatory break itself as source', () => {
            // 'ab\ncd': the first line drew two characters and consumed three.
            const [first, second] = breakIntoLines(mono('ab\ncd'), CELL * 40);

            expect(textOf(first!)).toBe('ab');
            expect(first!.endsAt).toBe(3);
            expect(second!.startsAt).toBe(3);
            expect(second!.endsAt).toBe(5);
        });

        it('counts CRLF as the TWO units it occupies', () => {
            // The pair collapses to one break but still takes two positions in
            // the text, and a caret map that assumed one would drift.
            const [first, second] = breakIntoLines(mono('ab\r\ncd'), CELL * 40);

            expect(first!.endsAt).toBe(4);
            expect(second!.startsAt).toBe(4);
            expect(second!.endsAt).toBe(6);
        });

        it('gives an empty line produced by a break a position of its own', () => {
            const lines = breakIntoLines(mono('a\n\nb'), CELL * 40);

            expect(lines.map(textOf)).toEqual(['a', '', 'b']);
            expect(lines[1]!.startsAt).toBe(2);
            expect(lines[1]!.endsAt).toBe(3);
        });

        it('measures offsets in UTF-16 units, as every text offset is', () => {
            // An astral character is ONE code point and TWO units. Iterating a
            // string yields code points, so counting steps would put every
            // offset after an emoji one short.
            const text = 'a\u{1F600}b cd';
            const [first] = breakIntoLines(mono(text), CELL * 40);

            expect(first!.endsAt).toBe(text.length);
            expect(text.slice(first!.startsAt, first!.endsAt)).toBe(text);
        });

        it('gives each piece its own start, even mid-WORD across a run boundary', () => {
            // One word whose second half is styled differently — a single
            // unbreakable segment holding two pieces. Two separate words would
            // be two segments, and each piece would start where its segment
            // did, which is the same answer for the wrong reason.
            const runs: StyledRun[] = [
                { text: 'Hel', font: MONO, sizePx: 16 },
                { text: 'lo there', font: SERIF, sizePx: 16 },
            ];

            const [line] = breakIntoLines(runs, CELL * 40);
            const combined = 'Hello there';

            expect(line!.pieces[0]!.sourceStart).toBe(0);
            expect(line!.pieces[1]!.sourceStart).toBe(3);
            for (const piece of line!.pieces) {
                expect(combined.slice(piece.sourceStart, piece.sourceStart + piece.text.length))
                    .toBe(piece.text);
            }
        });

        it('keeps trailing spaces inside the span, though not in the width', () => {
            // The space hangs into the margin for measurement and is still text
            // the caret can sit in.
            const [first] = breakIntoLines(mono('aaa bbb'), CELL * 6);

            expect(textOf(first!)).toBe('aaa ');
            expect(first!.endsAt).toBe(4);
        });
    });

        describe('leaders', () => {
        const tabPiece = (leader?: 'dot' | 'hyphen') => {
            const [line] = breakIntoLines([{ text: 'a\tb', font: MONO, sizePx: 16 }], 10000, {
                tabStops: {
                    stops: [{
                        positionPx: CELL * 10,
                        align: 'left',
                        ...(undefined === leader ? {} : { leader }),
                    }],
                    defaultPx: 1000,
                },
            });

            return line!.pieces[1]!;
        };

        it('records the leader its stop asked for on the TAB piece', () => {
            // Only the breaker knows which stop the tab landed on, and the
            // piece's own width is exactly the span the fill has to cover.
            expect(tabPiece('dot').leader).toBe('dot');
            expect(tabPiece('dot').widthPx).toBeCloseTo(CELL * 9, 6);
        });

        it('leaves an ordinary tab without one', () => {
            expect(tabPiece().leader).toBeUndefined();
        });
    });

    describe('comparing widths', () => {
        it('fits a line that is over by a rounding error', () => {
            // Twips and points reach this engine by different routes, so the
            // same length computed two ways differs in the last bits — a
            // writing area of `page - left - right` against a tab stop declared
            // at exactly that width is out by about 1e-13px. Compared exactly,
            // a right stop at the margin puts its page number on a line of its
            // own, which is every table of contents ever written.
            const width = CELL * 4;
            const [line] = breakIntoLines(
                [{ text: 'aaaa', font: MONO, sizePx: 16 }],
                width - 1e-13,
            );

            expect(line!.pieces.length).toBe(1);
            // And it is not reported as OVERFLOWING either: a word a
            // ten-trillionth of a pixel too wide has not overrun anything, and
            // saying so would let it be drawn past the margin on purpose.
            expect(line!.overflows).toBe(false);
        });

        it('still breaks a line that is genuinely too long', () => {
            // The tolerance is a ten-thousandth of a pixel, not a licence.
            const lines = breakIntoLines(
                [{ text: 'aaaa bbbb', font: MONO, sizePx: 16 }],
                CELL * 4 + 0.01,
            );

            expect(lines.length).toBe(2);
        });
    });

    describe('aligned stops', () => {
        const STOP = CELL * 20;

        /**
         * Where the text after the tab starts, measured from the line's left.
         *
         * Piece 0 is what precedes the tab and piece 1 is the tab itself, so
         * their widths together are the start of what follows.
         */
        const startOf = (
            text: string,
            align: 'center' | 'right' | 'decimal',
            decimalSymbol?: string,
        ): number => {
            const [line] = breakIntoLines([{ text, font: MONO, sizePx: 16 }], 10000, {
                tabStops: {
                    stops: [{ positionPx: STOP, align }],
                    defaultPx: 1000,
                    ...(undefined === decimalSymbol ? {} : { decimalSymbol }),
                },
            });

            return line!.pieces[0]!.widthPx + line!.pieces[1]!.widthPx;
        };

        it('ENDS the following text on a right stop', () => {
            expect(startOf('a\tAAAA', 'right')).toBeCloseTo(STOP - CELL * 4, 6);
        });

        it('straddles a centre stop', () => {
            expect(startOf('a\tAAAA', 'center')).toBeCloseTo(STOP - CELL * 2, 6);
        });

        it('puts the SEPARATOR on a decimal stop', () => {
            // Two digits before the point, so the point lands on the stop and
            // the digits after it hang past it.
            expect(startOf('a\t12.34', 'decimal')).toBeCloseTo(STOP - CELL * 2, 6);
        });

        it('lines up on the separator the DOCUMENT names', () => {
            // A comma across most of Europe: aligned on a full stop, a
            // comma-decimal document ranges its numbers on nothing at all.
            expect(startOf('a\t12,34', 'decimal', ',')).toBeCloseTo(STOP - CELL * 2, 6);
            // The same text against a full stop finds no separator, so all of
            // it is measured instead.
            expect(startOf('a\t12,34', 'decimal')).toBeCloseTo(STOP - CELL * 5, 6);
        });

        it('treats a number with no separator as all of it', () => {
            expect(startOf('a\t1234', 'decimal')).toBeCloseTo(STOP - CELL * 4, 6);
        });

        it('COLLAPSES when the text will not fit before the stop', () => {
            // Verified against LibreOffice: a right stop followed by more text
            // than fits leaves it immediately after what preceded the tab,
            // rather than dragging it back over that text.
            const [line] = breakIntoLines(
                [{ text: `a\t${'a'.repeat(40)}`, font: MONO, sizePx: 16 }],
                10000,
                { tabStops: { stops: [{ positionPx: CELL * 5, align: 'right' }], defaultPx: 1000 } },
            );

            expect(line!.pieces[1]!.widthPx).toBe(0);
        });

        it('ranges on the FIRST separator when a number holds two', () => {
            // A version string or a dotted date has several. Word ranges on the
            // first, and taking the last would put "1.2" left of the stop
            // instead of "1".
            expect(startOf('a\t1.2.3', 'decimal')).toBeCloseTo(STOP - CELL, 6);
        });

        it('leaves the repeating default stops LEFT, whatever the explicit ones are', () => {
            // Word offers no way to align a default stop. Inheriting the last
            // explicit stop's alignment would range text against a column the
            // document never asked for.
            const [line] = breakIntoLines([{ text: 'a\tb\tcccc', font: MONO, sizePx: 16 }], 10000, {
                tabStops: {
                    stops: [{ positionPx: CELL * 4, align: 'right' }],
                    defaultPx: CELL * 4,
                },
            });
            // First tab ranges 'b' to end at 4 cells; the second falls on a
            // default stop at 8 and simply advances to it.
            const upToSecondTab = line!.pieces
                .slice(0, 4)
                .reduce((sum, piece) => sum + piece.widthPx, 0);

            expect(upToSecondTab).toBeCloseTo(CELL * 8, 6);
        });

        it('measures only as far as the NEXT tab', () => {
            // The stop after this one takes over there, so text beyond it is
            // not what this stop is ranging.
            expect(startOf('a\tAAAA\tBBBBBBBB', 'right')).toBeCloseTo(STOP - CELL * 4, 6);
        });
    });

    describe('stretchable gaps', () => {
        it('counts the gaps BETWEEN words, not the spaces after the last one', () => {
            // Three words is two gaps. The run of spaces that ends a line is
            // trimmed off its width, so stretching it would push the last word
            // past the margin instead of filling the line.
            const [only] = breakIntoLines([{ text: 'aa bb cc   ', font: MONO, sizePx: 16 }], 1000);

            expect(only!.spaceGaps).toBe(2);
        });

        it('has no gap to stretch on a single word', () => {
            const [only] = breakIntoLines([{ text: 'aaaa', font: MONO, sizePx: 16 }], 1000);

            expect(only!.spaceGaps).toBe(0);
        });

        it('counts a RUN of spaces between two words as one gap', () => {
            const [only] = breakIntoLines([{ text: 'aa     bb', font: MONO, sizePx: 16 }], 1000);

            expect(only!.spaceGaps).toBe(1);
        });

        it('counts the gaps of each wrapped line separately', () => {
            // Two words per line at four cells: the gap that was consumed by the
            // wrap belongs to no line at all.
            const lines = breakIntoLines(
                [{ text: 'aa bb cc dd', font: MONO, sizePx: 16 }],
                CELL * 5,
            );

            expect(lines.map((each) => each.spaceGaps)).toEqual([1, 1]);
        });
    });

describe('tab stops', () => {
        /** The width a line reached, which for a tabbed line is the point. */
        const widthOf = (text: string, stops: object, columns = 40): number => {
            const [line] = breakIntoLines(mono(text), CELL * columns, { tabStops: stops });

            return line!.widthPx;
        };

        it('advances a tab to the next STOP, not by a glyph', () => {
            // The tab character maps to .notdef in every font here — 8.11px in
            // Carlito, 12.45px in Liberation Serif — and neither is a column.
            expect(widthOf('a\tb', { defaultPx: 100 })).toBeCloseTo(100 + CELL, 6);
        });

        it('advances to the NEXT stop each time, not by a fixed step', () => {
            // Two tabs from different starting points land in the same columns.
            expect(widthOf('a\tb\tc', { defaultPx: 100 })).toBeCloseTo(200 + CELL, 6);
            expect(widthOf('aaaaa\tb\tc', { defaultPx: 100 })).toBeCloseTo(200 + CELL, 6);
        });

        it('takes the explicit stops before the default ones', () => {
            expect(widthOf('a\tb', { stops: [{ positionPx: 30, align: 'left' }, { positionPx: 250, align: 'left' }], defaultPx: 100 }))
                .toBeCloseTo(30 + CELL, 6);
            expect(widthOf('a\tb\tc', { stops: [{ positionPx: 30, align: 'left' }, { positionPx: 250, align: 'left' }], defaultPx: 100 }))
                .toBeCloseTo(250 + CELL, 6);
        });

        it('moves PAST a stop the cursor is already sitting on', () => {
            // Three characters land exactly on the stop at three cells. A tab
            // there must go to the NEXT column — taking the one it is already at
            // would advance nothing, and two tabbed fields would print on top of
            // each other.
            const width = widthOf('aaa\tb', { stops: [{ positionPx: CELL * 3, align: 'left' }, { positionPx: CELL * 6, align: 'left' }], defaultPx: 100 });

            expect(width).toBeCloseTo(CELL * 7, 6);
        });

        it('falls to the COLUMN’s own default stops past the last explicit one', () => {
            // Otherwise a third tabbed column stops dead on the second one.
            //
            // The defaults are at multiples of the step from the margin — 100,
            // 200, 300 — and not a fresh repeat measured from the last explicit
            // stop, which would put this at 350. The pen reaches
            // 250 + CELL, so the first default past it is 300. The ones behind
            // 250 are simply out of reach, never having been removed.
            expect(widthOf('a\tb\tc\td', { stops: [{ positionPx: 30, align: 'left' }, { positionPx: 250, align: 'left' }], defaultPx: 100 }))
                .toBeCloseTo(300 + CELL, 6);
        });

        it('measures stops from the COLUMN, not from the line', () => {
            // A line starting 40 along its column — pushed there by an indent,
            // a first-line indent, or both — still has its tabs land in the
            // column's own places.
            const shifted = breakIntoLines(mono('a\tb'), CELL * 40, {
                tabStops: { defaultPx: 100, originOf: () => 40 },
            });

            // The stop at 100 is 60 from a line that starts 40 along.
            expect(shifted[0]!.widthPx).toBeCloseTo(60 + CELL, 6);
        });

        it('breaks the line when the tabbed text no longer fits', () => {
            // A tab that would overshoot the column has to wrap like anything
            // else, or the text runs off the page.
            const lines = breakIntoLines(mono('a\tb'), CELL * 5, { tabStops: { defaultPx: 1000 } });

            expect(lines.length).toBeGreaterThan(1);
        });

        it('leaves a tab in the source span, for a caret to land on', () => {
            const [line] = breakIntoLines(mono('ab\tcd'), CELL * 40, { tabStops: { defaultPx: 100 } });

            expect(line!.startsAt).toBe(0);
            expect(line!.endsAt).toBe(5);
        });

        it('does nothing when there are no stops at all', () => {
            // Better than an arbitrary glyph width: a caller that configures no
            // stops gets no phantom column.
            expect(widthOf('a\tb', {})).toBeCloseTo(CELL * 2, 6);
        });
    });
});

describe('a run that boxes itself', () => {
    const BORDER: BorderSide = {
        widthPx: 4, style: 'single' as BorderStyle, colorHex: '#123456', spacePx: 2,
    };

    it('carries `w:bdr` onto EVERY piece the run breaks into', () => {
        // Measured against LibreOffice: a boxed run broken over three lines
        // came out as three complete boxes, one per line. A piece that lost
        // the border on the way would be a line drawn without its box.
        const boxed: StyledRun[] = [{ text: 'aaa bbb ccc', font: MONO, sizePx: 16, border: BORDER }];
        const lines = breakIntoLines(boxed, CELL * 4);

        expect(lines.length).toBeGreaterThan(1);
        for (const line of lines) {
            for (const piece of line.pieces) {
                expect(piece.border).toBe(BORDER);
            }
        }
    });

    it('costs the line the room at BOTH ends of the run’s span', () => {
        // The box needs its room kept clear before the run and after it, so a
        // boxed run fits `2 × room` less on a line than the same words plain.
        // Charging one end, or none, or once a WORD, all change this number.
        const room = (BORDER.spacePx ?? 0) + BORDER.widthPx;
        const plain = breakIntoLines(mono('aaaa bbbb cccc'), CELL * 10);
        const boxed = breakIntoLines(
            [{ text: 'aaaa bbbb cccc', font: MONO, sizePx: 16, border: BORDER }],
            CELL * 10);

        // Wide enough for the plain words, and short by the room for the boxed.
        expect(plain[0]!.widthPx).toBeCloseTo(CELL * 9, 6);
        expect(boxed[0]!.widthPx).toBeCloseTo(CELL * 4 + 2 * room, 6);
        expect(boxed.length).toBeGreaterThan(plain.length);
    });

    it('charges the room again on every line the run carries onto', () => {
        // Each line's segment is its own closed box, so each keeps its own
        // room: three lines pay three times, not once between them.
        const room = (BORDER.spacePx ?? 0) + BORDER.widthPx;
        const boxed = breakIntoLines(
            [{ text: 'aaaa bbbb cccc', font: MONO, sizePx: 16, border: BORDER }],
            CELL * 10);

        for (const line of boxed) {
            expect(line.widthPx).toBeGreaterThan(2 * room);
        }
    });

    it('leaves the pieces of an unboxed run without one', () => {
        const lines = breakIntoLines(mono('aaa bbb'), CELL * 4);

        expect(lines.flatMap((line) => line.pieces).every((p) => undefined === p.border))
            .toBe(true);
    });

    it('boxes only the run that asked, not its neighbours on the line', () => {
        const mixed: StyledRun[] = [
            { text: 'aa', font: MONO, sizePx: 16 },
            { text: 'bb', font: MONO, sizePx: 16, border: BORDER },
            { text: 'cc', font: MONO, sizePx: 16 },
        ];
        const [line] = breakIntoLines(mixed, CELL * 20);

        expect(line!.pieces.map((piece) => undefined !== piece.border))
            .toEqual([false, true, false]);
    });
});
