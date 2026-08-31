/**
 * Find tests that assert geometry without saying where the number came from.
 *
 * This engine's rules are measured out of LibreOffice, one printed PDF at a
 * time. A test that pins a number without citing the measurement is pinning
 * whatever the code happened to do the day it was written — and three times in
 * one arc that turned out to be a defect the suite was protecting:
 *
 *   - a box's height taken from the LINE it shared instead of its own run
 *   - a split paragraph box left open, by analogy with a table row
 *   - every shared table edge drawn twice, asserted as `toBe(6)` under the
 *     comment "each of them really is drawn twice"
 *
 * The tell each time was a comment explaining why our answer is REASONABLE
 * rather than where the number came from. This looks for exactly that: a test
 * asserting a fractional coordinate, with no citation in the test or in the
 * `describe` around it.
 *
 * It is a prompt to go and measure, not a failure. Run it with:
 *
 *     npm run audit:claims
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join, relative } from 'node:path';
import { globSync } from 'node:fs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');

/**
 * Words that mean "this number was read off something", not invented.
 *
 * A citation is WORDS, never an identifier. This used to accept a tracker
 * id as evidence, which let a test cite something no reader of this package
 * can look up — and made the id load-bearing, so removing one silently
 * turned a cited test into an uncited one.
 */
const CITES = /LibreOffice|measured|probed|printed|the PDF|Word (?:draws|puts|does|centres|treats)|FONT FILE|file's own tables|OOXML schema/i;

/** A fractional coordinate in an assertion: the kind only a measurement gives. */
const GEOMETRY = /\.toBe(?:CloseTo)?\(\s*-?\d+\.\d+|toEqual\(\[\s*-?\d+\.\d+/;

const files = globSync('tests/**/*.spec.ts', { cwd: ROOT }).sort();
const flagged = [];

for (const file of files) {
    const text = readFileSync(join(ROOT, file), 'utf8');
    // Each `describe` carries context its tests lean on, so a citation there
    // counts for everything inside it.
    const blocks = text.split(/\n(?=\s*describe\()/);

    for (const block of blocks) {
        const preamble = block.slice(0, block.search(/\n\s*it\(/) + 1);
        const cited = CITES.test(preamble);

        for (const test of block.split(/\n(?=\s*it\()/).slice(1)) {
            const name = /^\s*it\('((?:[^'\\]|\\.)*)'/.exec(test)?.[1];
            if (undefined === name || !GEOMETRY.test(test) || cited || CITES.test(test)) {
                continue;
            }
            flagged.push(`${relative('.', file)}  ${name}`);
        }
    }
}

// The same question of the SOURCE: a constant is a claim about the world too,
// and `MINIMUM_RULE_PX` sat in the renderer for months as an invented 0.5px
// under a comment explaining why a floor was sensible.
// A number can also be a definition or a tolerance rather than a claim
// about the world. Those need saying so, but not measuring.
const SPEC = /defined by|the (?:ZIP|OOXML|OpenType|PDF) (?:format|spec)|signature|polynomial|magic|as OOXML defines|a tolerance, not a measurement/i;

for (const file of globSync('src/**/*.ts', { cwd: ROOT }).sort()) {
    const text = readFileSync(join(ROOT, file), 'utf8');

    for (const match of text.matchAll(/^(?:export )?const ([A-Z][A-Z_0-9]*) = (-?[\d.]+)/gm)) {
        // Whole small numbers are enumerations and table indices, not
        // measurements of anything: `LOOKUP_PAIR = 2` claims nothing.
        const value = Number(match[2]);
        if (Number.isInteger(value) && Math.abs(value) < 100) {
            continue;
        }

        const before = text.slice(0, match.index).split('\n').slice(-14).join('\n');
        if (CITES.test(before) || SPEC.test(before)) {
            continue;
        }
        flagged.push(`${relative('.', file)}  ${match[1]}`);
    }
}

if (0 === flagged.length) {
    console.log('Every pinned number says where it came from.');
} else {
    console.log(`${flagged.length} claim(s) pin a number with no measurement cited:\n`);
    for (const entry of flagged) {
        console.log(`  ${entry}`);
    }
    console.log('\nEach is a probe waiting to be run, not a failure.');
}
