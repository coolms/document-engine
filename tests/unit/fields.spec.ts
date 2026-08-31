import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import { fieldText, hasField, resolveFields } from '../../src/layout/fields.js';
import type { FieldKind } from '../../src/layout/line-breaker.js';
import type { NumeralStyle } from '../../src/text/numerals.js';
import { DEFAULT_PARAGRAPH_STYLE, type Block, type Paragraph } from '../../src/layout/page-layout.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MONO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Regular.ttf'))));

const CONTEXT = { pageNumber: 4, pageCount: 9 };

function paragraph(
    ...runs: { text: string; field?: FieldKind; format?: NumeralStyle }[]
): Paragraph {
    return {
        runs: runs.map((run) => ({
            text: run.text,
            font: MONO,
            sizePx: 16,
            ...(undefined === run.field ? {} : { field: run.field }),
            ...(undefined === run.format ? {} : { fieldFormat: run.format }),
        })),
        style: DEFAULT_PARAGRAPH_STYLE,
    };
}

const textOf = (block: Block): string =>
    'rows' in block ? '<table>' : block.runs.map((run) => run.text).join('');

describe('fieldText', () => {
    it('answers PAGE with the page and NUMPAGES with the count', () => {
        // Distinct numbers, because a fixture where they matched could not tell
        // one field from the other.
        expect(fieldText('page', CONTEXT)).toBe('4');
        expect(fieldText('numPages', CONTEXT)).toBe('9');
    });
});

describe('fieldText, written the way its switch asks', () => {
    it('writes the number in the style the field named', () => {
        expect(fieldText('page', CONTEXT, 'lowerRoman')).toBe('iv');
        expect(fieldText('numPages', CONTEXT, 'upperRoman')).toBe('IX');
    });

    it('falls back to digits when the field named nothing', () => {
        expect(fieldText('page', CONTEXT)).toBe('4');
    });
});

describe('hasField', () => {
    it('finds a field of any kind when no kind is named', () => {
        expect(hasField(paragraph({ text: '1', field: 'page' }))).toBe(true);
        expect(hasField(paragraph({ text: 'plain' }))).toBe(false);
    });

    it('distinguishes the kinds, because only NUMPAGES costs a second pass', () => {
        const page = paragraph({ text: '1', field: 'page' });

        expect(hasField(page, 'page')).toBe(true);
        expect(hasField(page, 'numPages')).toBe(false);
    });
});

describe('resolveFields', () => {
    it('writes each field run in ITS OWN style', () => {
        // Front matter numbered i, ii, iii beside a total in digits is an
        // ordinary thing for a document to ask for.
        const blocks = [paragraph(
            { text: '9', field: 'page', format: 'lowerRoman' },
            { text: ' of ' },
            { text: '9', field: 'numPages' },
        )];

        expect(resolveFields(blocks, CONTEXT).map(textOf)).toEqual(['iv of 9']);
    });

    it('replaces a field run and leaves the literal text beside it', () => {
        const blocks = [paragraph(
            { text: 'Page ' },
            { text: '99', field: 'page' },
            { text: ' of ' },
            { text: '77', field: 'numPages' },
        )];

        expect(resolveFields(blocks, CONTEXT).map(textOf)).toEqual(['Page 4 of 9']);
    });

    it('does not touch the blocks it was given', () => {
        // The same header is drawn on every page with a different number in it.
        // Resolving in place would make each page overwrite the one before.
        const blocks = [paragraph({ text: '99', field: 'page' })];

        resolveFields(blocks, CONTEXT);

        expect(textOf(blocks[0]!)).toBe('99');
    });

    it('returns the very same blocks when none of them has a field', () => {
        // Furniture is re-resolved for every page, and a header with no page
        // number is the common case — copying it per page would allocate a whole
        // tree for no change.
        const blocks = [paragraph({ text: 'Chapter One' })];

        expect(resolveFields(blocks, CONTEXT)).toBe(blocks);
    });

    it('leaves a table alone', () => {
        const table = {
            rows: [], columnWidthsPx: [], cellMarginLeftPx: 0, cellMarginRightPx: 0,
            cellMarginTopPx: 0, cellMarginBottomPx: 0, spaceBeforePx: 0, spaceAfterPx: 0,
            pageBreakBefore: false,
        };
        const blocks = [table, paragraph({ text: '99', field: 'page' })];

        const resolved = resolveFields(blocks, CONTEXT);

        expect(resolved[0]).toBe(table);
        expect(textOf(resolved[1]!)).toBe('4');
    });
});
