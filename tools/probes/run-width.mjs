/**
 * How wide a string is, in the face the probe is using.
 *
 * ## Why this exists
 *
 * A printed page gives absolute positions, so reading a feature's own geometry
 * out of one usually means subtracting the text before it: a box's left edge is
 * `where the run after it starts` less `how wide that run is`. Estimating that
 * second number by eye is how the same measurement came out 17.50,
 * 16.40 and finally 18.02 — three readings of one page,
 * two of them published. `P-before` is 34.43 and not "about 36", and nothing
 * about the page ever changed.
 *
 * So: measure the run you are subtracting, or subtract nothing. This prints the
 * engine's own advance for a string, which is the number the print agrees with
 * to a hundredth for every face this repository ships.
 *
 *   node tools/probes/run-width.mjs 10 "P-before" "L-" "alpha beta gamma"
 *
 * The first argument is the size in POINTS, the rest are strings. The face is
 * Liberation Serif, which every hand-built probe in `tests/fixtures/docx` uses;
 * pass `--font "Carlito"` for another one the manifest carries.
 */

import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

import { FontCatalogue } from '../../dist/word/font-catalogue.js';
import { advanceOf } from '../../dist/layout/line-breaker.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const FONT_DIR = join(HERE, '../../assets/fonts');

const argv = process.argv.slice(2);
let family = 'Liberation Serif';
const flag = argv.indexOf('--font');
if (flag >= 0) {
    family = argv[flag + 1];
    argv.splice(flag, 2);
}

const [sizeArgument, ...strings] = argv;
const sizePt = Number(sizeArgument);
if (!Number.isFinite(sizePt) || 0 === strings.length) {
    console.error('usage: run-width.mjs <sizePt> <string> [string …] [--font "Family"]');
    process.exit(2);
}

const manifest = JSON.parse(readFileSync(join(FONT_DIR, 'fonts.manifest.json'), 'utf8'));
const fonts = FontCatalogue.load(manifest, (name) => new Uint8Array(readFileSync(join(FONT_DIR, name))));
const resolved = fonts.resolve(family, false, false);

if (null !== resolved.substitutedFor) {
    console.error(`!! "${family}" is not in the manifest; measuring in ${resolved.font.familyName}`);
}

// Through `advanceOf`, so the answer is the one the line breaker would use —
// including the rule that a run is UNKERNED unless the document said `w:kern`.
// A probe measuring kerned widths would disagree with the page.
const sizePx = sizePt * 96 / 72;
for (const text of strings) {
    const px = advanceOf(text, { font: resolved.font, sizePx });
    console.log(`${(px * 72 / 96).toFixed(2).padStart(8)}pt  ${JSON.stringify(text)}`);
}
