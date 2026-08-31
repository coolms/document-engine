import { FontReader, type TableRecord } from './sfnt.js';

/**
 * Pair kerning read out of the OpenType GPOS table.
 *
 * ## Why this exists
 *
 * Advance widths alone are not how wide text is. Renderers adjust the space
 * between specific pairs — `AV` tucks together, `To` tucks under the T.
 * Measured against Chrome, our unkerned sum of `AVAVAVAV` came out **14.4375px
 * too wide at 16px** (about 462 font units per pair), which is enough to move a
 * page boundary.
 *
 * ## But a document decides WHETHER, and it is usually no
 *
 * This class answers how much; it does not answer whether. That belongs to the
 * run's `w:kern`, and Chrome is the wrong renderer to have asked: a browser
 * kerns by default and a WORD PROCESSOR does not. Printed by LibreOffice,
 * `AVAVAVAVAV` with no `w:kern` came out **11.60pt WIDER** than the same string
 * with one — the wider number being the plain sum of advances.
 *
 * So every caller measures through {@link advanceOf}, which kerns only where
 * the document asked. Nothing in the fixture corpus asks, including both real
 * documents and everything this platform generates.
 *
 * ## Why GPOS rather than the `kern` table
 *
 * Of the five families we ship, only Liberation Sans and Serif carry the legacy
 * `kern` table at all. Carlito — our DEFAULT font — Caladea and Liberation Mono
 * put kerning in GPOS exclusively. Reading `kern` would have covered two
 * families and silently under-measured the other three.
 *
 * ## Scope
 *
 * Pair adjustment (lookup type 2), both formats, plus the extension wrapper
 * (type 9) that real fonts use to reach past the 16-bit offset limit. Only the
 * horizontal advance of the FIRST glyph is applied: that is what changes line
 * width. Placement, vertical adjustment and device tables move where a glyph is
 * DRAWN without changing how far the pen travels, so they are irrelevant to
 * measurement and are deliberately not read.
 *
 * Contextual kerning (types 7/8) is not handled. It is rare in text fonts and
 * would need the full shaping pipeline; leaving it out is a known, bounded
 * inaccuracy rather than a silent one.
 */
export class GposKerning {
    private readonly cache = new Map<number, number>();

    private constructor(private readonly subtables: readonly PairSubtable[]) {
    }

    /**
     * Kerning for a font that has none.
     *
     * A real instance rather than a null: callers measure the same way whether
     * or not the font kerns, and the alternative — parsing a zero offset as if
     * it were a GPOS table — would read the sfnt header and invent adjustments
     * out of it.
     */
    static none(): GposKerning {
        return new GposKerning([]);
    }

    /** True when the font contributed no usable pair kerning. */
    get isEmpty(): boolean {
        return 0 === this.subtables.length;
    }

    /**
     * Horizontal adjustment between two glyphs, in font units. Negative pulls
     * them together, which is the common case.
     */
    between(leftGlyph: number, rightGlyph: number): number {
        const key = leftGlyph * 0x10000 + rightGlyph;
        const cached = this.cache.get(key);
        if (cached !== undefined) {
            return cached;
        }

        let adjustment = 0;
        for (const subtable of this.subtables) {
            const value = subtable.lookup(leftGlyph, rightGlyph);
            if (0 !== value) {
                // First subtable to express an opinion wins: lookups are
                // ordered and an earlier one takes precedence.
                adjustment = value;
                break;
            }
        }

        this.cache.set(key, adjustment);

        return adjustment;
    }

    /**
     * Parse the GPOS table, collecting every pair-adjustment subtable reachable
     * from a `kern` feature.
     *
     * Features are gathered across ALL scripts rather than resolved per
     * script/language. Latin and Cyrillic share the same pair adjustments in
     * these fonts, and picking a script would mean knowing the run's language
     * before measuring it — a dependency that buys nothing here and would have
     * to be threaded through the whole engine.
     */
    static parse(reader: FontReader, gpos: TableRecord): GposKerning {
        const base = gpos.offset;
        const featureListOffset = base + reader.uint16(base + 6);
        const lookupListOffset = base + reader.uint16(base + 8);

        const wanted = new Set<number>();
        const featureCount = reader.uint16(featureListOffset);
        for (let i = 0; i < featureCount; i++) {
            const record = featureListOffset + 2 + i * 6;
            if ('kern' !== reader.tag(record)) {
                continue;
            }

            const feature = featureListOffset + reader.uint16(record + 4);
            const lookupCount = reader.uint16(feature + 2);
            for (let j = 0; j < lookupCount; j++) {
                wanted.add(reader.uint16(feature + 4 + j * 2));
            }
        }

        const subtables: PairSubtable[] = [];
        const lookupCount = reader.uint16(lookupListOffset);
        for (const index of wanted) {
            if (index >= lookupCount) {
                continue;
            }
            const lookup = lookupListOffset + reader.uint16(lookupListOffset + 2 + index * 2);
            collectPairSubtables(reader, lookup, subtables);
        }

        return new GposKerning(subtables);
    }
}

interface PairSubtable {
    lookup(left: number, right: number): number;
}

const LOOKUP_PAIR = 2;
const LOOKUP_EXTENSION = 9;

function collectPairSubtables(reader: FontReader, lookup: number, out: PairSubtable[]): void {
    const lookupType = reader.uint16(lookup);
    const subtableCount = reader.uint16(lookup + 4);

    for (let i = 0; i < subtableCount; i++) {
        const subtable = lookup + reader.uint16(lookup + 6 + i * 2);

        if (LOOKUP_EXTENSION === lookupType) {
            // An extension subtable is a redirect: it exists because a 16-bit
            // offset cannot reach the whole table in a large font. Ignoring it
            // would silently drop most of a big font's kerning.
            const extensionType = reader.uint16(subtable + 2);
            const target = subtable + reader.uint32(subtable + 4);
            if (LOOKUP_PAIR === extensionType) {
                readPairSubtable(reader, target, out);
            }

            continue;
        }

        if (LOOKUP_PAIR === lookupType) {
            readPairSubtable(reader, subtable, out);
        }
    }
}

function readPairSubtable(reader: FontReader, subtable: number, out: PairSubtable[]): void {
    const format = reader.uint16(subtable);
    const valueFormat1 = reader.uint16(subtable + 4);
    const valueFormat2 = reader.uint16(subtable + 6);

    // Where XAdvance sits inside value1, and how long the whole pair of value
    // records is. A font that adjusts something other than advance contributes
    // nothing to width, so it is skipped rather than mis-read.
    const xAdvanceOffset = valueFieldOffset(valueFormat1, VALUE_X_ADVANCE);
    if (null === xAdvanceOffset) {
        return;
    }
    const recordSize = valueRecordSize(valueFormat1) + valueRecordSize(valueFormat2);

    if (1 === format) {
        out.push(readPairFormat1(reader, subtable, xAdvanceOffset, recordSize));

        return;
    }
    if (2 === format) {
        out.push(readPairFormat2(reader, subtable, xAdvanceOffset, recordSize));
    }
}

/**
 * Format 1: an explicit list of second glyphs per first glyph.
 *
 * Flattened into a map at parse time. These lists are small (a few thousand
 * pairs in a text font) and the alternative — a binary search per character
 * pair per measurement — would run inside the line-breaking loop.
 */
function readPairFormat1(
    reader: FontReader,
    subtable: number,
    xAdvanceOffset: number,
    recordSize: number,
): PairSubtable {
    const coverage = readCoverage(reader, subtable + reader.uint16(subtable + 2));
    const pairSetCount = reader.uint16(subtable + 8);
    const pairs = new Map<number, number>();

    for (let i = 0; i < pairSetCount; i++) {
        const firstGlyph = coverage[i];
        if (firstGlyph === undefined) {
            continue;
        }

        const pairSet = subtable + reader.uint16(subtable + 10 + i * 2);
        const pairValueCount = reader.uint16(pairSet);

        for (let j = 0; j < pairValueCount; j++) {
            // Each record is secondGlyph(2) followed by the two value records.
            const record = pairSet + 2 + j * (2 + recordSize);
            const secondGlyph = reader.uint16(record);
            const adjustment = reader.int16(record + 2 + xAdvanceOffset);

            if (0 !== adjustment) {
                pairs.set(firstGlyph * 0x10000 + secondGlyph, adjustment);
            }
        }
    }

    return {
        lookup: (left, right) => pairs.get(left * 0x10000 + right) ?? 0,
    };
}

/**
 * Format 2: glyphs grouped into classes, with a class-by-class matrix.
 *
 * Evaluated on demand rather than flattened — the matrix is class1 x class2 and
 * expanding it to every glyph pair would be enormous for no benefit, since only
 * the pairs a document actually contains are ever asked for.
 */
function readPairFormat2(
    reader: FontReader,
    subtable: number,
    xAdvanceOffset: number,
    recordSize: number,
): PairSubtable {
    const coverage = new Set(readCoverage(reader, subtable + reader.uint16(subtable + 2)));
    const classDef1 = readClassDef(reader, subtable + reader.uint16(subtable + 8));
    const classDef2 = readClassDef(reader, subtable + reader.uint16(subtable + 10));
    const class1Count = reader.uint16(subtable + 12);
    const class2Count = reader.uint16(subtable + 14);
    const matrix = subtable + 16;

    return {
        lookup: (left, right) => {
            // Coverage gates the FIRST glyph only; a glyph outside it is not
            // kerned by this subtable however its class is defined.
            if (!coverage.has(left)) {
                return 0;
            }

            const class1 = classDef1(left);
            const class2 = classDef2(right);
            if (class1 >= class1Count || class2 >= class2Count) {
                return 0;
            }

            const record = matrix + (class1 * class2Count + class2) * recordSize;

            return reader.int16(record + xAdvanceOffset);
        },
    };
}

/** Coverage table → the glyphs it covers, in coverage-index order. */
function readCoverage(reader: FontReader, offset: number): number[] {
    const format = reader.uint16(offset);
    const glyphs: number[] = [];

    if (1 === format) {
        const count = reader.uint16(offset + 2);
        for (let i = 0; i < count; i++) {
            glyphs.push(reader.uint16(offset + 4 + i * 2));
        }

        return glyphs;
    }

    if (2 === format) {
        const rangeCount = reader.uint16(offset + 2);
        for (let i = 0; i < rangeCount; i++) {
            const range = offset + 4 + i * 6;
            const start = reader.uint16(range);
            const end = reader.uint16(range + 2);
            const startIndex = reader.uint16(range + 4);

            for (let glyph = start; glyph <= end; glyph++) {
                glyphs[startIndex + (glyph - start)] = glyph;
            }
        }

        return glyphs;
    }

    throw new Error(`Unknown coverage format ${format} in GPOS.`);
}

/** ClassDef table → a function giving a glyph's class (0 when unlisted). */
function readClassDef(reader: FontReader, offset: number): (glyph: number) => number {
    const format = reader.uint16(offset);

    if (1 === format) {
        const startGlyph = reader.uint16(offset + 2);
        const glyphCount = reader.uint16(offset + 4);

        return (glyph) => {
            const index = glyph - startGlyph;

            return index >= 0 && index < glyphCount ? reader.uint16(offset + 6 + index * 2) : 0;
        };
    }

    if (2 === format) {
        const rangeCount = reader.uint16(offset + 2);
        const ranges: { start: number; end: number; classId: number }[] = [];
        for (let i = 0; i < rangeCount; i++) {
            const range = offset + 4 + i * 6;
            ranges.push({
                start: reader.uint16(range),
                end: reader.uint16(range + 2),
                classId: reader.uint16(range + 4),
            });
        }

        return (glyph) => {
            for (const range of ranges) {
                if (glyph >= range.start && glyph <= range.end) {
                    return range.classId;
                }
            }

            return 0;
        };
    }

    throw new Error(`Unknown class definition format ${format} in GPOS.`);
}

const VALUE_X_ADVANCE = 0x0004;

/** ValueRecord fields are present in bit order; each present field is 2 bytes. */
function valueRecordSize(valueFormat: number): number {
    let size = 0;
    for (let bit = 1; bit <= 0x0080; bit <<= 1) {
        if (0 !== (valueFormat & bit)) {
            size += 2;
        }
    }

    return size;
}

/** Byte offset of one field inside a ValueRecord, or null when absent. */
function valueFieldOffset(valueFormat: number, field: number): number | null {
    if (0 === (valueFormat & field)) {
        return null;
    }

    let offset = 0;
    for (let bit = 1; bit < field; bit <<= 1) {
        if (0 !== (valueFormat & bit)) {
            offset += 2;
        }
    }

    return offset;
}
