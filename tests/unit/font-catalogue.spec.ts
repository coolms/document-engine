import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { FontCatalogue, fontStyleKey, offeredFontFamilies, type FontManifest } from '../../src/word/font-catalogue.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MANIFEST = JSON.parse(readFileSync(join(FONT_DIR, 'fonts.manifest.json'), 'utf8')) as FontManifest;

function catalogue(): FontCatalogue {
    return FontCatalogue.load(MANIFEST, (file) => new Uint8Array(readFileSync(join(FONT_DIR, file))));
}

describe('FontCatalogue', () => {
    describe('substitution', () => {
        it('maps the Microsoft faces a document names to the ones we ship', () => {
            // This is the mapping the whole vendored font set exists for. Each
            // pair is metric-compatible, which is what keeps an inherited
            // document from repaginating when it is opened here.
            const fonts = catalogue();

            expect(fonts.resolve('Times New Roman', false, false).family).toBe('Liberation Serif');
            expect(fonts.resolve('Arial', false, false).family).toBe('Liberation Sans');
            expect(fonts.resolve('Courier New', false, false).family).toBe('Liberation Mono');
            expect(fonts.resolve('Calibri', false, false).family).toBe('Carlito');
            expect(fonts.resolve('Cambria', false, false).family).toBe('Caladea');
        });

        it('matches a family name whatever its case', () => {
            // Documents in the wild spell it "ARIAL", "Arial" and "arial".
            // A case-sensitive lookup silently falls through to the default,
            // which is a different face with different metrics.
            const fonts = catalogue();

            expect(fonts.resolve('ARIAL', false, false).family).toBe('Liberation Sans');
            expect(fonts.resolve('arial', false, false).family).toBe('Liberation Sans');
            expect(fonts.resolve('  Arial  ', false, false).family).toBe('Liberation Sans');
        });

        it('says WHAT it substituted, so the swap can be reported', () => {
            const fonts = catalogue();

            expect(fonts.resolve('Arial', false, false).substitutedFor).toBe('Arial');
            // Naming the family we actually hold is not a substitution.
            expect(fonts.resolve('Liberation Sans', false, false).substitutedFor).toBeNull();
            expect(fonts.resolve(null, false, false).substitutedFor).toBeNull();
        });

        it('falls back to the default family rather than failing', () => {
            // A document naming a font nobody has must still lay out.
            const fonts = catalogue();
            const resolved = fonts.resolve('Papyrus', false, false);

            expect(resolved.family).toBe(fonts.defaultFamily);
            expect(resolved.substitutedFor).toBe('Papyrus');
        });

        it('knows which names it can honour', () => {
            const fonts = catalogue();

            expect(fonts.knows('Calibri')).toBe(true);
            expect(fonts.knows('calibri')).toBe(true);
            expect(fonts.knows('Papyrus')).toBe(false);
        });
    });

    describe('weight and slope', () => {
        it('picks a different FILE for each of the four combinations', () => {
            const fonts = catalogue();
            const faces = [
                fonts.resolve('Calibri', false, false).font,
                fonts.resolve('Calibri', true, false).font,
                fonts.resolve('Calibri', false, true).font,
                fonts.resolve('Calibri', true, true).font,
            ];

            // Four distinct parsed fonts, not one face reused. `fontSynthesis`
            // is none precisely because a faux bold has different advances than
            // the real bold face.
            expect(new Set(faces).size).toBe(4);
        });

        it('measures bold WIDER than regular, as a real bold face does', () => {
            const fonts = catalogue();
            const regular = fonts.resolve('Calibri', false, false).font;
            const bold = fonts.resolve('Calibri', true, false).font;

            expect(bold.measureAdvance('Wagon flight', 16).widthPt)
                .toBeGreaterThan(regular.measureAdvance('Wagon flight', 16).widthPt);
        });

        it('names the four style keys the manifest uses', () => {
            expect(fontStyleKey(false, false)).toBe('regular');
            expect(fontStyleKey(true, false)).toBe('bold');
            expect(fontStyleKey(false, true)).toBe('italic');
            expect(fontStyleKey(true, true)).toBe('boldItalic');
        });
    });

    describe('loading', () => {
        it('parses each file once, however often it is asked for', () => {
            // A document repeats the same few faces on every run; re-parsing a
            // 400KB font per run would dominate the cost of laying it out.
            let reads = 0;
            const fonts = FontCatalogue.load(MANIFEST, (file) => {
                reads++;

                return new Uint8Array(readFileSync(join(FONT_DIR, file)));
            });

            const first = fonts.resolve('Arial', false, false).font;
            const second = fonts.resolve('Arial', false, false).font;
            const alias = fonts.resolve('Helvetica', false, false).font;

            expect(reads).toBe(1);
            expect(second).toBe(first);
            expect(alias).toBe(first);
        });

        it('does not load a single font until one is asked for', () => {
            let reads = 0;
            FontCatalogue.load(MANIFEST, (file) => {
                reads++;

                return new Uint8Array(readFileSync(join(FONT_DIR, file)));
            });

            expect(reads).toBe(0);
        });

        it('takes its defaults from the manifest, not from a constant', () => {
            // The manifest is read by the PHP render side too. A default
            // hard-coded here would drift from the one used there.
            const fonts = catalogue();

            expect(fonts.defaultFamily).toBe(MANIFEST.defaults.family);
            expect(fonts.defaultSizePt).toBe(MANIFEST.defaults.sizePt);
            expect(fonts.knows(fonts.defaultFamily)).toBe(true);
        });

        it('refuses a manifest whose default family it does not contain', () => {
            // Otherwise every unknown font resolves to nothing at all, and the
            // failure appears as a crash somewhere in measurement.
            const broken: FontManifest = {
                families: MANIFEST.families,
                defaults: { family: 'Nonexistent', sizePt: 11 },
            };
            const fonts = FontCatalogue.load(broken, (file) =>
                new Uint8Array(readFileSync(join(FONT_DIR, file))));

            expect(() => fonts.resolve('Papyrus', false, false)).toThrow(/no family "Nonexistent"/);
        });
    });
});

describe('offeredFontFamilies', () => {
    it('offers exactly one name per vendored family, default first', () => {
        // The list an editor shows. It is DERIVED so that installing a family
        // is a manifest entry and four files -- the literal it replaced offered
        // Georgia, Verdana and Tahoma, which nothing here could measure.
        expect(offeredFontFamilies(MANIFEST)).toEqual([
            'Calibri', 'Times New Roman', 'Arial', 'Courier New', 'Georgia', 'Cambria',
        ]);
    });

    it('offers a name the catalogue can actually resolve', () => {
        // The guard that matters: an offered name the catalogue does not know
        // is measured as the base face while the browser paints something else,
        // which is the whole defect. Asserted over whatever the manifest holds
        // TODAY, so a family added later is covered without editing this.
        const fonts = catalogue();

        for (const name of offeredFontFamilies(MANIFEST)) {
            expect(fonts.knows(name), name).toBe(true);
            expect(fonts.resolve(name, false, false).substitutedFor).toBe(name);
        }
    });

    it('offers every family the manifest ships', () => {
        // The other direction: Cambria was vendored for months and never
        // offered, so an author could not choose the one face we had.
        expect(offeredFontFamilies(MANIFEST)).toHaveLength(Object.keys(MANIFEST.families).length);
    });

    it('names a family by its FIRST substitute, which is what a document writes', () => {
        // `Calibri`, never `Carlito`: the name goes into the .docx unchanged
        // and has to be one Word and LibreOffice resolve.
        const manifest: FontManifest = {
            families: {
                Vendored: { substitutes: ['Asked For', 'Vendored'], files: {} },
            },
            defaults: { family: 'Vendored', sizePt: 11 },
        };

        expect(offeredFontFamilies(manifest)).toEqual(['Asked For']);
    });

    it('falls back to the family name when an entry lists no substitutes', () => {
        const manifest: FontManifest = {
            families: { Lonely: { substitutes: [], files: {} } },
            defaults: { family: 'Lonely', sizePt: 11 },
        };

        expect(offeredFontFamilies(manifest)).toEqual(['Lonely']);
    });

    it('shows a name once even when two families answer to it', () => {
        // A select with the same option twice is a bug the manifest cannot see.
        const manifest: FontManifest = {
            families: {
                First: { substitutes: ['Shared'], files: {} },
                Second: { substitutes: ['Shared'], files: {} },
            },
            defaults: { family: 'First', sizePt: 11 },
        };

        expect(offeredFontFamilies(manifest)).toEqual(['Shared']);
    });
});
