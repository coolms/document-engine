import { describe, expect, it } from 'vitest';

import { Numbering, NumberingCounters } from '../../src/word/numbering.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** One abstract numbering with the levels given, plus `instances` numIds. */
function numbering(levels: string, instances: readonly [string, string][] = [['1', '7']]): Numbering {
    const nums = instances
        .map(([numId, abstractId]) => `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`)
        .join('');

    return Numbering.parse(
        `<w:numbering ${W}><w:abstractNum w:abstractNumId="7">${levels}</w:abstractNum>${nums}</w:numbering>`,
    );
}

function level(ilvl: number, format: string, text: string, extra = ''): string {
    return `<w:lvl w:ilvl="${ilvl}"><w:numFmt w:val="${format}"/>`
        + `<w:lvlText w:val="${text}"/>${extra}</w:lvl>`;
}

/** Markers for a run of paragraphs at the given levels. */
function markers(sheet: Numbering, steps: readonly (readonly [string, number])[]): string[] {
    const counters = new NumberingCounters();

    return steps.map(([numId, ilvl]) => counters.next(sheet, numId, ilvl));
}

describe('Numbering', () => {
    describe('the two levels of indirection', () => {
        it('resolves a numId through its abstract definition', () => {
            const sheet = numbering(level(0, 'decimal', '%1.'));

            expect(sheet.level('1', 0)?.format).toBe('decimal');
            expect(sheet.level('1', 0)?.text).toBe('%1.');
        });

        it('counts two numIds sharing one abstract definition TOGETHER', () => {
            // Measured against LibreOffice, and the opposite of what
            // this test asserted before it was: three numIds on one abstract,
            // used in turn, counted straight through it. A change of numId is
            // not a new list — an OVERRIDE is, which is why Word writes one
            // whenever a list is asked to begin again.
            const sheet = numbering(level(0, 'decimal', '%1.'), [['1', '7'], ['2', '7']]);

            expect(markers(sheet, [['1', 0], ['2', 0], ['1', 0], ['2', 0]]))
                .toEqual(['1.', '2.', '3.', '4.']);
        });

        it('counts two ABSTRACTS separately', () => {
            // The other side of it: two definitions are two lists, and nothing
            // here should have made them share.
            const sheet = Numbering.parse(
                `<w:numbering ${W}>`
                + `<w:abstractNum w:abstractNumId="7">${level(0, 'decimal', '%1.')}</w:abstractNum>`
                + `<w:abstractNum w:abstractNumId="8">${level(0, 'decimal', '%1.')}</w:abstractNum>`
                + '<w:num w:numId="1"><w:abstractNumId w:val="7"/></w:num>'
                + '<w:num w:numId="2"><w:abstractNumId w:val="8"/></w:num>'
                + '</w:numbering>',
            );

            expect(markers(sheet, [['1', 0], ['2', 0], ['1', 0], ['2', 0]]))
                .toEqual(['1.', '1.', '2.', '2.']);
        });

        it('knows nothing about a numId the file does not define', () => {
            expect(numbering(level(0, 'decimal', '%1.')).level('99', 0)).toBeNull();
            expect(numbering(level(0, 'decimal', '%1.')).level('1', 9)).toBeNull();
        });

        it('has no lists at all when there is no numbering part', () => {
            expect(Numbering.empty().level('1', 0)).toBeNull();
        });
    });

    describe('bullets', () => {
        it('draws the level\'s glyph and counts nothing', () => {
            const sheet = numbering(level(0, 'bullet', '●'));

            expect(markers(sheet, [['1', 0], ['1', 0], ['1', 0]])).toEqual(['●', '●', '●']);
        });

        it('does not COUNT, which a deeper level\'s pattern would reveal', () => {
            // Level one is bullets and level two names it with %1. However many
            // bullets have gone by, the pattern shows the level's start — a
            // bullet has no number for anything to refer to.
            const sheet = numbering(
                level(0, 'bullet', '\u25cf') + level(1, 'decimal', '%1.%2'),
            );

            expect(markers(sheet, [
                ['1', 0], ['1', 0], ['1', 0], ['1', 1],
            ])).toEqual(['\u25cf', '\u25cf', '\u25cf', '1.1']);
        });

        it('reads the font a bullet glyph is drawn in', () => {
            // Symbol and Wingdings are what Word reaches for, and their glyphs
            // are nothing like the same code points in a text face.
            const sheet = numbering(level(0, 'bullet', '',
                '<w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr>'));

            expect(sheet.level('1', 0)?.fontFamily).toBe('Symbol');
        });
    });

    describe('counting', () => {
        it('counts from one and keeps going', () => {
            const sheet = numbering(level(0, 'decimal', '%1.'));

            expect(markers(sheet, [['1', 0], ['1', 0], ['1', 0]])).toEqual(['1.', '2.', '3.']);
        });

        it('starts where w:start says', () => {
            const sheet = numbering(level(0, 'decimal', '%1.', '<w:start w:val="5"/>'));

            expect(markers(sheet, [['1', 0], ['1', 0]])).toEqual(['5.', '6.']);
        });

        it('RESETS a deeper level when a shallower one advances', () => {
            // Without the reset the second section's sub-items continue from the
            // first section's count: 1.1, 1.2, then 2.3 instead of 2.1.
            const sheet = numbering(level(0, 'decimal', '%1.') + level(1, 'decimal', '%1.%2'));

            expect(markers(sheet, [
                ['1', 0], ['1', 1], ['1', 1],
                ['1', 0], ['1', 1],
            ])).toEqual(['1.', '1.1', '1.2', '2.', '2.1']);
        });

        it('does not reset a SHALLOWER level when a deeper one advances', () => {
            const sheet = numbering(level(0, 'decimal', '%1.') + level(1, 'decimal', '%1.%2'));

            expect(markers(sheet, [['1', 0], ['1', 1], ['1', 0]]))
                .toEqual(['1.', '1.1', '2.']);
        });

        it('substitutes each level\'s own format into the pattern', () => {
            // "%1" refers to LEVEL ONE, which may be counted differently from
            // the level doing the referring.
            const sheet = numbering(
                level(0, 'upperLetter', '%1.') + level(1, 'decimal', '%1.%2)'),
            );

            expect(markers(sheet, [['1', 0], ['1', 1], ['1', 1]]))
                .toEqual(['A.', 'A.1)', 'A.2)']);
        });
    });

    describe('number formats', () => {
        const of = (format: string, count: number): string[] => {
            const sheet = numbering(level(0, format, '%1'));

            return markers(sheet, Array.from({ length: count }, () => ['1', 0] as const));
        };

        it('counts in lower and upper letters, REPEATING past z', () => {
            // Word's sequence is a…z then aa, bb, cc — not the spreadsheet's
            // aa, ab, ac.
            const lower = of('lowerLetter', 28);

            expect(lower.slice(0, 3)).toEqual(['a', 'b', 'c']);
            expect(lower[25]).toBe('z');
            expect(lower[26]).toBe('aa');
            expect(lower[27]).toBe('bb');
            expect(of('upperLetter', 2)).toEqual(['A', 'B']);
        });

        it('counts in roman numerals, both cases', () => {
            expect(of('lowerRoman', 9)).toEqual(['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix']);
            expect(of('upperRoman', 4).slice(3)).toEqual(['IV']);
        });

        it('builds the awkward roman numerals correctly', () => {
            const sheet = numbering(level(0, 'lowerRoman', '%1', '<w:start w:val="1900"/>'));

            expect(markers(sheet, [['1', 0]])).toEqual(['mcm']);
        });

        it('falls back to decimal for a format it does not draw', () => {
            // Wrong glyphs beat a list that silently loses its numbers.
            expect(of('cardinalText', 2)).toEqual(['1', '2']);
        });
    });

    describe('restarting', () => {
        it('keeps counting when w:lvlRestart is zero', () => {
            // How a document numbers its figures 1..40 straight through
            // chapters that each restart their own sections.
            const sheet = numbering(
                level(0, 'decimal', '%1.')
                + level(1, 'decimal', 'Figure %2', '<w:lvlRestart w:val="0"/>'),
            );

            expect(markers(sheet, [
                ['1', 0], ['1', 1], ['1', 1],
                ['1', 0], ['1', 1],
            ])).toEqual(['1.', 'Figure 1', 'Figure 2', '2.', 'Figure 3']);
        });

        it('restarts by default, which is the common case', () => {
            const sheet = numbering(level(0, 'decimal', '%1.') + level(1, 'decimal', 'Figure %2'));

            expect(markers(sheet, [['1', 0], ['1', 1], ['1', 0], ['1', 1]]))
                .toEqual(['1.', 'Figure 1', '2.', 'Figure 1']);
        });
    });

    describe('the marker suffix', () => {
        it('reads w:suff, and defaults to a tab as Word does', () => {
            expect(numbering(level(0, 'decimal', '%1.')).level('1', 0)?.suffix).toBe('tab');
            expect(numbering(level(0, 'decimal', '%1.', '<w:suff w:val="space"/>'))
                .level('1', 0)?.suffix).toBe('space');
            expect(numbering(level(0, 'decimal', '%1.', '<w:suff w:val="nothing"/>'))
                .level('1', 0)?.suffix).toBe('nothing');
            // An unknown value is a tab rather than a crash.
            expect(numbering(level(0, 'decimal', '%1.', '<w:suff w:val="wat"/>'))
                .level('1', 0)?.suffix).toBe('tab');
        });
    });

    describe('indents', () => {
        it('reads the level\'s left and hanging indents', () => {
            const sheet = numbering(level(0, 'bullet', '●',
                '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'));

            expect(sheet.level('1', 0)?.indentLeftTwips).toBe(720);
            expect(sheet.level('1', 0)?.indentHangingTwips).toBe(360);
        });

        it('accepts w:start as a spelling of w:left', () => {
            const sheet = numbering(level(0, 'bullet', '●',
                '<w:pPr><w:ind w:start="720"/></w:pPr>'));

            expect(sheet.level('1', 0)?.indentLeftTwips).toBe(720);
        });
    });

    describe('w:lvlOverride', () => {
        /**
         * One decimal abstract, and numIds on it with the overrides given.
         *
         * Every list here shares a definition on purpose: with the count living
         * on the abstract, an override is the ONLY thing that can make two of
         * them differ, so the difference cannot come from anywhere else.
         */
        const overridden = (...instances: readonly string[]): Numbering => Numbering.parse(
            `<w:numbering ${W}><w:abstractNum w:abstractNumId="7">`
            + `${level(0, 'decimal', '%1.', '<w:start w:val="1"/>')}</w:abstractNum>`
            + instances.map((override, index) =>
                `<w:num w:numId="${index + 1}"><w:abstractNumId w:val="7"/>${override}</w:num>`)
                .join('')
            + '</w:numbering>',
        );

        const START_AT_SEVEN = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride>';
        const AS_LETTERS = '<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0">'
            + '<w:start w:val="3"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="(%1)"/>'
            + '</w:lvl></w:lvlOverride>';

        it('restarts the count at w:startOverride', () => {
            // Measured: a plain list ran 1. 2. 3. and the next numId,
            // on the SAME abstract, printed 7. 8. 9. — so the override reaches
            // a counter that is already running and sets it.
            const sheet = overridden('', START_AT_SEVEN);

            expect(markers(sheet, [['1', 0], ['1', 0], ['1', 0], ['2', 0], ['2', 0], ['2', 0]]))
                .toEqual(['1.', '2.', '3.', '7.', '8.', '9.']);
        });

        it('spends the override ONCE, at the first paragraph that uses it', () => {
            // The measured order, interleaved: the overridden list resumed at
            // 11. and 12. rather than restarting at its stated 7. A reset on
            // every use would have printed 7. 8. there, which is how a document
            // ends up with two items numbered the same.
            const sheet = overridden('', START_AT_SEVEN);

            expect(markers(sheet, [
                ['1', 0], ['1', 0],   // 1. 2.
                ['2', 0], ['2', 0],   // the override fires: 7. 8.
                ['1', 0], ['1', 0],   // 9. 10.
                ['2', 0], ['2', 0],   // spent, so 11. 12.
            ])).toEqual(['1.', '2.', '7.', '8.', '9.', '10.', '11.', '12.']);
        });

        it('replaces the level outright where the override states a w:lvl', () => {
            // Format and pattern both come from the replacement — measured as
            // `(C)` where the abstract would have said `3.`.
            const sheet = overridden(AS_LETTERS);

            expect(sheet.level('1', 0)?.format).toBe('upperLetter');
            expect(sheet.level('1', 0)?.text).toBe('(%1)');
            expect(markers(sheet, [['1', 0], ['1', 0]])).toEqual(['(C)', '(D)']);
        });

        it('reads a replacement level\'s w:start as a value, not a restart', () => {
            // The sharpest number in the probe. The same instance printed `(C)`
            // from a `w:start` of 3 with the counter fresh, and `(J)` — the
            // tenth letter — when the list before it had run to 9. A `w:lvl`
            // inside an override says what the level IS; only `w:startOverride`
            // says where the count begins again.
            const sheet = overridden('', START_AT_SEVEN, AS_LETTERS);

            expect(markers(sheet, [
                ['1', 0], ['1', 0], ['1', 0],   // 1. 2. 3.
                ['2', 0], ['2', 0], ['2', 0],   // 7. 8. 9.
                ['3', 0], ['3', 0], ['3', 0],
            ])).toEqual(['1.', '2.', '3.', '7.', '8.', '9.', '(J)', '(K)', '(L)']);
        });

        it('ignores an override that does not say WHICH level it is for', () => {
            // `w:ilvl` is required on `w:lvlOverride`, so a file without it is
            // malformed and there is nothing measured to follow. Reading it as
            // level zero — the tempting guess — would restart a list the
            // document never asked to restart, so the definition's own
            // numbering stands and nothing is invented.
            const sheet = overridden('<w:lvlOverride><w:startOverride w:val="7"/></w:lvlOverride>');

            expect(markers(sheet, [['1', 0], ['1', 0]])).toEqual(['1.', '2.']);
        });

        it('leaves an override for ANOTHER level alone', () => {
            // `w:ilvl` on the override says which level it speaks for, and a
            // list whose first level is overridden must not have its second
            // level restarted too.
            const sheet = overridden(
                '<w:lvlOverride w:ilvl="3"><w:startOverride w:val="7"/></w:lvlOverride>');

            expect(markers(sheet, [['1', 0], ['1', 0]])).toEqual(['1.', '2.']);
        });
    });
});
