import { formatNumeral, type NumeralStyle } from '../text/numerals.js';
import type { FieldKind, StyledRun } from './line-breaker.js';
import type { Block, Paragraph } from './page-layout.js';

/**
 * Resolving `PAGE` and `NUMPAGES` against the page being drawn.
 *
 * Its own module for the same reason `alignment` is: `page-layout` needs it at
 * runtime and so would anything else that draws furniture, while the types it
 * works on live in `page-layout`. Importing only TYPES from there keeps the two
 * off each other's runtime graph.
 */

export interface FieldContext {
    /** The printed number of the page being drawn, counting from one. */
    readonly pageNumber: number;
    /** How many pages the finished document has. */
    readonly pageCount: number;
    /**
     * `w:pgNumType/@w:fmt` — how this SECTION writes its page numbers.
     *
     * A field's own `\*` switch wins where it has one; this is the answer for
     * the many that do not, and it is how a document actually asks for roman
     * front matter. Measured against LibreOffice: a section saying
     * `lowerRoman` printed i and ii, and the next section's page printed 1.
     *
     * It answers for the page NUMBER only. `NUMPAGES` is a count of pages
     * rather than one page's name, and nothing measured says a section's
     * format should reach it.
     */
    readonly pageNumberFormat?: NumeralStyle;
}

/**
 * The text a field resolves to, written the way its switch asks for.
 *
 * `PAGE \* roman` on front matter is the case that matters: pages numbered
 * i, ii, iii rather than 1, 2, 3.
 */
export function fieldText(
    kind: FieldKind,
    context: FieldContext,
    format?: NumeralStyle,
): string {
    const stated = format ?? ('page' === kind ? context.pageNumberFormat : undefined);

    return formatNumeral(
        'page' === kind ? context.pageNumber : context.pageCount,
        stated ?? 'decimal',
    );
}

/**
 * Replace every field run in these blocks with what it resolves to.
 *
 * Returns new blocks and leaves the originals alone: the same header is drawn
 * on every page with a different number in it, so mutating would make each page
 * overwrite the one before.
 *
 * Blocks with no fields in them are returned AS THEY ARE, not copied. Furniture
 * is re-resolved for every page of the document, and a header with no page
 * number in it is the common case — copying it each time would allocate a fresh
 * tree per page for no change at all.
 */
export function resolveFields(blocks: readonly Block[], context: FieldContext): readonly Block[] {
    if (!blocks.some((block) => hasField(block))) {
        return blocks;
    }

    return blocks.map((block) => isParagraph(block) ? resolveParagraph(block, context) : block);
}

/**
 * Whether anything here is a field — of a given kind, if one is named.
 *
 * The kind matters: `PAGE` is answered exactly on the first pass, because a
 * page knows its own number. `NUMPAGES` is the only one worth laying a document
 * out twice for.
 */
export function hasField(block: Block, kind?: FieldKind): boolean {
    return isParagraph(block)
        && block.runs.some((run) => undefined !== run.field
            && (undefined === kind || kind === run.field));
}

function resolveParagraph(paragraph: Paragraph, context: FieldContext): Paragraph {
    if (!paragraph.runs.some((run) => undefined !== run.field)) {
        return paragraph;
    }

    return {
        ...paragraph,
        runs: paragraph.runs.map((run: StyledRun) => undefined === run.field
            ? run
            : { ...run, text: fieldText(run.field, context, run.fieldFormat) }),
    };
}

/**
 * A block is a paragraph when it is not a table.
 *
 * Spelled out here rather than imported from `page-layout`, which would put
 * this module on that one's runtime graph — the very thing keeping the two
 * apart. The test is the same one `isTable` makes.
 */
function isParagraph(block: Block): block is Paragraph {
    return !('rows' in block);
}
