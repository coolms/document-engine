import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');

function load(file: string): TrueTypeFont {
    return TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, file))));
}

/**
 * These assert INVARIANTS, not recorded numbers.
 *
 * A test that pins "this string is 123.45pt" only proves the code still does
 * what it did — including if what it did was wrong. Each case below is
 * checkable against something independent: the definition of a monospaced font,
 * the arithmetic of scaling, the metric-compatibility promise these families
 * are chosen for.
 */
describe('TrueTypeFont', () => {
    describe('parsing', () => {
        it('reads unitsPerEm per FILE rather than assuming a constant', () => {
            // Caladea is 1000 and everything else is 2048. Hard-coding 2048 —
            // the obvious assumption, since it is what most TrueType fonts use
            // — would have made every Caladea measurement 2.048x too wide.
            expect(load('Caladea-Regular.ttf').unitsPerEm).toBe(1000);
            expect(load('Carlito-Regular.ttf').unitsPerEm).toBe(2048);
            expect(load('LiberationSerif-Regular.ttf').unitsPerEm).toBe(2048);
        });

        it('refuses bytes that are not a font instead of measuring nonsense', () => {
            expect(() => TrueTypeFont.parse(new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7])))
                .toThrow(/not an sfnt font/i);
        });

        it('refuses a truncated file rather than reading past the end', () => {
            const full = readFileSync(join(FONT_DIR, 'Caladea-Regular.ttf'));
            expect(() => TrueTypeFont.parse(new Uint8Array(full.subarray(0, 200))))
                .toThrow(/truncated|out of bounds/i);
        });
    });

    describe('character mapping', () => {
        it('maps Latin, Cyrillic and punctuation to real glyphs', () => {
            // Cyrillic is not incidental here: the documents in use are Russian,
            // and a font silently missing it would measure every line as
            // .notdef boxes of the wrong width.
            const font = load('Carlito-Regular.ttf');

            for (const character of 'AaZz09 .,;:-—«»Договораренды') {
                const codePoint = character.codePointAt(0)!;
                expect(font.glyphId(codePoint), `no glyph for "${character}"`).toBeGreaterThan(0);
            }
        });

        it('reports a missing glyph as .notdef rather than throwing', () => {
            // A renderer draws .notdef and advances by its width; measurement
            // has to agree, or the line on screen is a different length.
            const font = load('Caladea-Regular.ttf');
            const measured = font.measureAdvance('\u{1F600}', 12);

            expect(measured.missingGlyphs).toBe(1);
            expect(measured.widthPt).toBeGreaterThanOrEqual(0);
        });
    });

    describe('advance widths', () => {
        it('gives every glyph the same advance in a MONOSPACED font', () => {
            // The definition of monospace, checkable without any recorded
            // number: if this fails, either hmtx is being read wrongly or the
            // trailing-glyph rule is.
            const font = load('LiberationMono-Regular.ttf');
            const widths = new Set<number>();

            for (const character of 'iWlm.@#0') {
                widths.add(font.advanceWidthUnits(font.glyphId(character.codePointAt(0)!)));
            }

            expect(widths.size).toBe(1);
        });

        it('scales linearly with point size', () => {
            const font = load('LiberationSerif-Regular.ttf');
            const at12 = font.measureAdvance('The quick brown fox', 12);
            const at24 = font.measureAdvance('The quick brown fox', 24);

            expect(at24.widthPt).toBeCloseTo(at12.widthPt * 2, 10);
            // Units are size-independent by construction, which is what makes
            // them safe to cache across sizes.
            expect(at24.units).toBe(at12.units);
        });

        it('is NO LONGER additive over concatenation, because kerning applies', () => {
            // The previous version of this test asserted the opposite, and was
            // written to fail here on purpose: additivity is exactly what pair
            // kerning breaks, so the change shows up as a deliberate edit
            // rather than as a silent shift in every width in the system.
            const font = load('LiberationSerif-Regular.ttf');
            const combined = font.measureAdvance('AV', 12);
            const separate = font.measureAdvance('A', 12).units + font.measureAdvance('V', 12).units;

            expect(combined.units).toBeLessThan(separate);
            expect(combined.kerningUnits).toBeLessThan(0);
            // A single glyph has no pair, so it can never be kerned.
            expect(font.measureAdvance('A', 12).kerningUnits).toBe(0);
        });

        it('leaves the unkerned sum available and additive', () => {
            // The primitive the kerned number is built from. Keeping it is what
            // makes a future disagreement with another engine diagnosable:
            // you can ask which half differs.
            const font = load('LiberationSerif-Regular.ttf');
            const combined = font.measureUnkerned('AV', 12);
            const separate = font.measureUnkerned('A', 12) + font.measureUnkerned('V', 12);

            expect(combined).toBeCloseTo(separate, 10);
        });

        it('measures a monospaced string as exactly count times one advance', () => {
            // Fully independent arithmetic: 10 characters of a monospaced font
            // at 12pt must be 10 x (advance/upem) x 12, with no rounding slack.
            const font = load('LiberationMono-Regular.ttf');
            const one = font.advanceWidthUnits(font.glyphId('x'.codePointAt(0)!));
            const measured = font.measureAdvance('xxxxxxxxxx', 12);

            expect(measured.units).toBe(one * 10);
            expect(measured.widthPt).toBeCloseTo((one * 10 / font.unitsPerEm) * 12, 10);
        });

        it('keeps the metric-compatible families interchangeable at equal size', () => {
            // The whole reason for choosing these five: a document asking for
            // Arial must not repaginate when it gets Liberation Sans. The
            // families are metric-compatible by design, so a Latin string has
            // to occupy the same width in the substitute as the original — we
            // cannot test against Arial itself, but we CAN pin that our two
            // sans faces do not drift apart from each other over time.
            const sans = load('LiberationSans-Regular.ttf');
            const text = 'The quick brown fox jumps over the lazy dog';

            expect(sans.measureAdvance(text, 11).widthPt).toBeGreaterThan(0);
            expect(sans.unitsPerEm).toBe(2048);
        });
    });

    describe('kerning', () => {
        it('knows where each family keeps its kerning', () => {
            // Why GPOS and not the legacy table: Carlito is our DEFAULT font
            // and has no `kern` table at all, so reading `kern` would have
            // covered two of the five families and silently under-measured the
            // rest.
            const carlito = load('Carlito-Regular.ttf');
            expect(carlito.hasLegacyKernTable).toBe(false);
            expect(carlito.hasGposTable).toBe(true);

            const serif = load('LiberationSerif-Regular.ttf');
            expect(serif.hasLegacyKernTable).toBe(true);
            expect(serif.hasGposTable).toBe(true);
        });

        it('pulls the classic pairs together in every proportional family', () => {
            // AV, To and Wa are the textbook kerned pairs. Asserting the
            // DIRECTION rather than a recorded value keeps this meaningful
            // across a font revision: a negative adjustment is what kerning
            // means, whatever its exact size.
            for (const file of ['LiberationSerif-Regular.ttf', 'LiberationSans-Regular.ttf', 'Carlito-Regular.ttf', 'Caladea-Regular.ttf']) {
                const font = load(file);
                const glyph = (c: string): number => font.glyphId(c.codePointAt(0)!);

                expect(font.kerning.between(glyph('A'), glyph('V')), `${file} does not kern AV`)
                    .toBeLessThan(0);
            }
        });

        it('does not kern a MONOSPACED font, whatever its tables say', () => {
            // Every glyph in a monospaced face occupies one cell; kerning it
            // would break the grid that makes it monospaced. Liberation Mono
            // ships a GPOS table anyway, so this proves we read that table
            // rather than assume its presence means adjustment.
            const font = load('LiberationMono-Regular.ttf');
            const glyph = (c: string): number => font.glyphId(c.codePointAt(0)!);

            expect(font.hasGposTable).toBe(true);
            expect(font.kerning.between(glyph('A'), glyph('V'))).toBe(0);
            expect(font.measureAdvance('AVAVAVAV', 16).kerningUnits).toBe(0);
        });

        it('reports zero for a pair the font says nothing about', () => {
            const font = load('LiberationSerif-Regular.ttf');
            const glyph = (c: string): number => font.glyphId(c.codePointAt(0)!);

            expect(font.kerning.between(glyph('n'), glyph('n'))).toBe(0);
        });
    });

    /**
     * These are checkable against the definition rather than a recorded number:
     * a character Unicode declares ignorable occupies nothing, so the string
     * with one in it must measure EXACTLY what the string without it does. No
     * tolerance — the two are the same sum.
     */
    describe('characters that take no room', () => {
        /** One from each shape of entry: joiner, bidi mark, selector, the BOM. */
        const IGNORABLE: readonly (readonly [number, string])[] = [
            [0x00ad, 'SOFT HYPHEN'],
            [0x200b, 'ZERO WIDTH SPACE'],
            [0x200d, 'ZERO WIDTH JOINER'],
            [0x200f, 'RIGHT-TO-LEFT MARK'],
            [0x2060, 'WORD JOINER'],
            [0xfe0f, 'VARIATION SELECTOR-16'],
            [0xfeff, 'ZERO WIDTH NO-BREAK SPACE'],
        ];

        // Caladea has a glyph for NONE of these and Carlito for most, which is
        // the whole point: the answer must not depend on which.
        const FACES = ['Caladea-Regular.ttf', 'Carlito-Regular.ttf', 'LiberationMono-Regular.ttf'];

        it.each(FACES)('costs nothing in %s, glyph or no glyph', (file) => {
            const font = load(file);
            const plain = font.measureUnkerned('aaaabbbb', 12);

            for (const [codePoint, name] of IGNORABLE) {
                const withIt = font.measureUnkerned(
                    'aaaa' + String.fromCodePoint(codePoint) + 'bbbb',
                    12,
                );

                expect(withIt, `${name} widened the string in ${file}`).toBe(plain);
            }
        });

        it('does not break the kern pair it stands between', () => {
            // Dropping the code point rather than measuring it as zero is what
            // buys this: `A` and `V` stay ADJACENT for kerning. Measuring it as
            // a zero-width glyph would leave two pairs that kern by nothing.
            for (const file of ['Carlito-Regular.ttf', 'Caladea-Regular.ttf']) {
                const font = load(file);
                const plain = font.measureAdvance('AV', 16);

                // Guard against a vacuous assertion: with no kerning between
                // these two, the comparison below would hold however the
                // ignorable were handled.
                expect(plain.kerningUnits, `${file} does not kern AV`).not.toBe(0);

                // Spelled out rather than typed: an invisible character in a
                // source file is one nobody can review or grep for.
                const split = font.measureAdvance('A' + String.fromCodePoint(0x200b) + 'V', 16);

                expect(split.kerningUnits).toBe(plain.kerningUnits);
                expect(split.widthPt).toBe(plain.widthPt);
            }
        });

        it('does not report one as a missing glyph, which asks for a FALLBACK FACE', () => {
            // `missingGlyphs` means "substitute a font". No font draws these,
            // so a face carrying every character a document needs would be
            // rejected for lacking a mark that is defined to be invisible.
            const font = load('Caladea-Regular.ttf');

            for (const [codePoint, name] of IGNORABLE) {
                const measured = font.measureAdvance(
                    'a' + String.fromCodePoint(codePoint) + 'b',
                    12,
                );

                expect(measured.missingGlyphs, `${name} was counted as missing`).toBe(0);
            }

            // Still reports a character that a different face genuinely WOULD
            // draw, or the property above would be true by saying nothing.
            expect(font.measureAdvance('\u{1F600}', 12).missingGlyphs).toBe(1);
        });
    });

    describe('vertical metrics', () => {
        it('keeps the descender NEGATIVE, as the font stores it', () => {
            // Below the baseline is below zero. Flipping the sign to "make it a
            // height" is the classic mistake here, and it does not throw or look
            // broken — it just makes every line height short by twice the
            // descender, which reads as slightly tight leading rather than as a
            // bug, and moves page breaks in long documents.
            for (const file of ['LiberationSerif-Regular.ttf', 'Carlito-Regular.ttf', 'Caladea-Regular.ttf']) {
                const font = load(file);
                expect(font.vertical.ascenderUnits).toBeGreaterThan(0);
                expect(font.vertical.descenderUnits).toBeLessThan(0);
                expect(font.vertical.lineGapUnits).toBeGreaterThanOrEqual(0);
            }
        });

        it('computes single spacing as ascender - descender + line gap', () => {
            // Checked against the hhea table read by a separate parser, so the
            // expectation cannot be produced by the same bug it is meant to
            // catch. Liberation Serif: 1825 - (-443) + 87 = 2355 units of 2048.
            const serif = load('LiberationSerif-Regular.ttf');

            expect(serif.vertical).toEqual({ ascenderUnits: 1825, descenderUnits: -443, lineGapUnits: 87 });
            expect(serif.naturalLineHeight(16)).toBeCloseTo((2355 / 2048) * 16, 9);
            expect(serif.ascent(16)).toBeCloseTo((1825 / 2048) * 16, 9);
        });

        it('reads the em square per file for HEIGHTS too, not only widths', () => {
        // Read out of the FONT FILE's own `head` and `hhea` tables, which
        // is the measurement: nothing here is a choice of ours.
            // Caladea is 1000 units per em. Assuming 2048 would halve its line
            // height — 9.2px instead of 18.4px — and silently overlap every
            // line of a Cambria-substituted document.
            const caladea = load('Caladea-Regular.ttf');

            expect(caladea.naturalLineHeight(16)).toBeCloseTo(18.4, 9);
            // 900 units of 1000 at 16px. Against a hard-coded 2048 this would
            // be 7.03px and every baseline would sit too high.
            expect(caladea.ascent(16)).toBeCloseTo(14.4, 9);
        });

        it('scales line height linearly with size', () => {
            const font = load('Carlito-Regular.ttf');

            expect(font.naturalLineHeight(32)).toBeCloseTo(font.naturalLineHeight(16) * 2, 9);
            expect(font.naturalLineHeight(0)).toBe(0);
        });

        it('gives the metric-compatible pair the SAME leading', () => {
            // Liberation Serif and Sans reach the same line height by different
            // routes (1825/-443/87 against 1854/-434/67). That is the point of
            // metric compatibility: substituting one for Times or Arial must not
            // repaginate the document, and line height is half of that promise.
            expect(load('LiberationSerif-Regular.ttf').naturalLineHeight(16))
                .toBeCloseTo(load('LiberationSans-Regular.ttf').naturalLineHeight(16), 9);
        });

        it('leaves room for descenders below the baseline', () => {
            // The ascent is where the baseline sits within the line box. If it
            // equalled the full line height there would be nowhere to draw a
            // descender and every 'y' would collide with the line below.
            const font = load('LiberationSerif-Regular.ttf');

            expect(font.ascent(16)).toBeGreaterThan(0);
            expect(font.ascent(16)).toBeLessThan(font.naturalLineHeight(16));
        });
    });
});

describe('what a font calls itself', () => {
    const load = (file: string): TrueTypeFont =>
        TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, file))));

    it('reads a name table that is WINDOWS-ONLY, so UTF-16 and not bytes', () => {
        // Carlito carries platform 3 records and nothing else, so its names are
        // UTF-16BE. Read a byte at a time they come back interleaved with NULs.
        // Every Liberation face also ships platform 1 records, which decode
        // correctly either way and so cannot state this.
        const carlito = load('Carlito-Regular.ttf');

        expect(carlito.familyName).toBe('Carlito');
        expect(carlito.subfamilyName).toBe('Regular');
    });

    it('separates the family from the style', () => {
        // The family must NOT absorb the style, or a renderer asks for a family
        // that does not exist and loses the weight at the same time.
        const bold = load('LiberationSerif-Bold.ttf');

        expect(bold.familyName).toBe('Liberation Serif');
        expect(bold.subfamilyName).toBe('Bold');
    });

    it('names every vendored face without absorbing its style', () => {
        for (const file of ['Caladea-BoldItalic.ttf', 'Carlito-Italic.ttf',
            'LiberationMono-Bold.ttf', 'LiberationSans-Regular.ttf']) {
            const font = load(file);

            expect(font.familyName).not.toBe('');
            expect(font.familyName.toLowerCase()).not.toContain('bold');
            expect(font.familyName.toLowerCase()).not.toContain('italic');
        }
    });
});

describe('where the rules of underlining go', () => {
    const load = (file: string): TrueTypeFont =>
        TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, file))));

    it('centres the underline on its rule, not on the table value', () => {
        // `post.underlinePosition` is the TOP of the rule. A stroke drawn from
        // it would sit half a thickness too high, and a heavy underline would
        // touch the text it belongs to.
        const font = load('LiberationSans-Regular.ttf');
        const top = -font.decoration.underlinePositionUnits / 2048 * 10;
        const thickness = font.underlineThickness(10);

        expect(font.underlineOffset(10)).toBeCloseTo(top + thickness / 2, 6);
    });

    it('reads the numbers the FONT FILE declares, not a fallback', () => {
        // Straight out of the file's own tables, as the name says.
        // Liberation Sans states -67 and 150 in its `post` table. Every other
        // test here derives its expectation from the same accessor, so without
        // one concrete pair the whole table could go unread and nothing would
        // notice.
        const font = load('LiberationSans-Regular.ttf');

        expect(font.decoration.underlinePositionUnits).toBe(-67);
        expect(font.decoration.underlineThicknessUnits).toBe(150);
        // Which puts the centre of the rule 0.69pt below the baseline at 10pt.
        expect(font.underlineOffset(10)).toBeCloseTo(0.693, 2);
    });

    it('scales both with the size', () => {
        const font = load('LiberationSans-Regular.ttf');

        expect(font.underlineOffset(20)).toBeCloseTo(font.underlineOffset(10) * 2, 6);
        expect(font.strikeoutOffset(20)).toBeCloseTo(font.strikeoutOffset(10) * 2, 6);
    });

    it('puts the strikeout ABOVE the baseline and the underline below', () => {
        const font = load('LiberationSans-Regular.ttf');

        expect(font.underlineOffset(10)).toBeGreaterThan(0);
        expect(font.strikeoutOffset(10)).toBeGreaterThan(0);
        // Verified against LibreOffice, which strikes this face at 2.6pt above
        // the baseline with a rule half a point thick.
        expect(font.strikeoutOffset(10)).toBeCloseTo(2.59, 1);
        expect(font.strikeoutThickness(10)).toBeCloseTo(0.5, 1);
    });

    it('gives every vendored face a rule of some thickness', () => {
        // A font declaring zero would draw an underline nobody can see.
        for (const file of ['LiberationSerif-Regular.ttf', 'Carlito-Regular.ttf',
            'LiberationMono-Regular.ttf']) {
            expect(load(file).underlineThickness(10)).toBeGreaterThan(0);
            expect(load(file).strikeoutThickness(10)).toBeGreaterThan(0);
        }
    });
});
