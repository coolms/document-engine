import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import { PARITY_CASES } from './browser-parity.cases.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const FONT_DIR = join(HERE, '../../assets/fonts');
const OUT_DIR = join(HERE, '../../.conformance');

/**
 * Computes what the ENGINE thinks each parity case measures, and writes it
 * where the browser-side check can pick it up.
 *
 * Emitting a file rather than hard-coding numbers into the browser check keeps
 * one source of truth: the fonts. A recorded constant would drift the moment a
 * font is re-vendored, and it would drift silently — the check would keep
 * passing against a stale expectation, which is worse than no check.
 */
describe('parity expectations', () => {
    it('computes a width for every case and writes the fixture', () => {
        const results = PARITY_CASES.map((testCase) => {
            const bytes = new Uint8Array(readFileSync(join(FONT_DIR, testCase.file)));
            const font = TrueTypeFont.parse(bytes);
            const measured = font.measureAdvance(testCase.text, testCase.sizePx);

            expect(measured.missingGlyphs, `${testCase.file} lacks glyphs for "${testCase.text}"`).toBe(0);
            expect(measured.widthPt).toBeGreaterThan(0);

            return {
                ...testCase,
                unitsPerEm: font.unitsPerEm,
                units: measured.units,
                unkernedUnits: measured.unkernedUnits,
                kerningUnits: measured.kerningUnits,
                // sizePx is in CSS pixels; the arithmetic is unit-agnostic, so
                // the "points" the measurer returns are pixels here.
                expectedPx: measured.widthPt,
            };
        });

        mkdirSync(OUT_DIR, { recursive: true });
        writeFileSync(join(OUT_DIR, 'parity-expectations.json'), JSON.stringify(results, null, 2));

        expect(results).toHaveLength(PARITY_CASES.length);
    });
});
