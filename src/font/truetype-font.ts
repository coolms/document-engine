import { GposKerning } from './gpos-kerning.js';
import { FontReader, readTableDirectory, requireTable, type TableRecord, type TableTag } from './sfnt.js';

/**
 * A parsed font, reduced to what text measurement needs: how wide each
 * character is, and in what units.
 *
 * ## What this is and is not
 *
 * It answers "how far does the pen advance after drawing this character",
 * which is the only question line breaking asks. It does not read outlines,
 * shape complex scripts, or apply substitutions — those change what is DRAWN,
 * not (for Latin and Cyrillic) how far the pen moves.
 *
 *  **Kerning is not applied yet.** See {@link measureAdvance}. That is a
 * named gap with a test asserting it, not an oversight — pretending otherwise
 * would put a number in the engine that quietly disagrees with the renderer.
 */
/**
 * Characters that LOOK like another, for a font that carries only the other.
 *
 * `.notdef` is not a free answer: a renderer advances by its width, and in
 * Liberation Serif that is 0.7778 em where a hyphen is 0.3330 — more than twice
 * as wide. Measured: a word joined by a `w:noBreakHyphen` broke a
 * character earlier here than in the print, because U+2011 is absent from every
 * face this engine ships and was being measured as `.notdef`. LibreOffice drew
 * it 3.30pt wide, which is the ordinary hyphen.
 *
 * Deliberately tiny, and only for characters that are the SAME MARK with
 * different behaviour — a non-breaking hyphen is a hyphen that refuses a break,
 * and its shape is not in question. A general "nearest glyph" table would be
 * inventing.
 *
 * The `0 === id` guard below — a font that HAS the character draws its own
 * glyph — cannot be tested while no face this engine ships carries U+2011,
 * which is the only entry here. A mutation removing it survives, honestly.
 */
const SAME_SHAPE = new Map<number, number>([
    [0x2011, 0x002d], // NON-BREAKING HYPHEN -> HYPHEN-MINUS
]);

/**
 * Unicode's `Default_Ignorable_Code_Point`s, as sorted inclusive ranges.
 *
 * The formatting marks — joiners, bidi controls, variation selectors, the soft
 * hyphen — that steer shaping and never carry ink. A font is free to leave them
 * out of `cmap`, and most do: then {@link TrueTypeFont.glyphId} answers 0 and
 * `.notdef` charges its box for a character DEFINED to occupy nothing.
 *
 * Measured through the faces this engine ships: Caladea has no glyph for a
 * single one of them and was charging 7.40px each at 12px, Carlito none for
 * U+2060 and was charging 6.08px. A word with a zero-width space between every
 * letter — what a soft-wrap hinter emits — measured 125% too wide in Caladea,
 * which moves the break and therefore the page.
 *
 * Not a hand-picked keep-list that could go stale: it is the Unicode property
 * itself, which is why the odd-looking entries (Hangul fillers, the musical
 * controls) are here rather than only the ones a Latin document hits.
 *
 * The property, and not merely a fallback for a missing glyph. Measured against
 * LibreOffice 26.2 through these same faces: a right-aligned `aaaaXbbbb` began
 * at exactly the x of `aaaabbbb` for every X below, in Caladea (which has none
 * of the glyphs) and in Carlito (which has most of them) alike. So the width is
 * zero whether or not the font carries the character, and skipping the code
 * point outright — rather than substituting a zero — is also what lets a kern
 * pair span one, which is what being ignorable in processing means.
 *
 *  MEASUREMENT only. These characters stay in the text handed to a renderer,
 * because U+200D joins an emoji sequence and U+FE0F picks its presentation:
 * both belong to the host that shapes, and both are destroyed by dropping them
 * here. This engine does not shape, so it must not decide they are redundant.
 */
const DEFAULT_IGNORABLE: readonly (readonly [number, number])[] = [
    [0x00ad, 0x00ad], // SOFT HYPHEN
    [0x034f, 0x034f], // COMBINING GRAPHEME JOINER
    [0x061c, 0x061c], // ARABIC LETTER MARK
    [0x115f, 0x1160], // HANGUL CHOSEONG/JUNGSEONG FILLER
    [0x17b4, 0x17b5], // KHMER INHERENT VOWELS
    [0x180b, 0x180f], // MONGOLIAN VARIATION SELECTORS, FREE VARIATION SELECTOR
    [0x200b, 0x200f], // ZWSP, ZWNJ, ZWJ, LRM, RLM
    [0x202a, 0x202e], // BIDI EMBEDDING AND OVERRIDE
    [0x2060, 0x206f], // WORD JOINER, INVISIBLE OPERATORS, BIDI ISOLATES, DEPRECATED FORMATS
    [0x3164, 0x3164], // HANGUL FILLER
    [0xfe00, 0xfe0f], // VARIATION SELECTOR-1..16
    [0xfeff, 0xfeff], // ZERO WIDTH NO-BREAK SPACE / BOM
    [0xffa0, 0xffa0], // HALFWIDTH HANGUL FILLER
    [0xfff0, 0xfff8], // unassigned, but reserved as ignorable
    [0x1bca0, 0x1bca3], // SHORTHAND FORMAT CONTROLS
    [0x1d173, 0x1d17a], // MUSICAL SYMBOL BEAM/SLUR/PHRASE CONTROLS
    [0xe0000, 0xe0fff], // TAGS and VARIATION SELECTORS SUPPLEMENT
];

/**
 * Whether a code point takes no room at all, whatever the font says.
 *
 * See {@link DEFAULT_IGNORABLE}. The early return is not a micro-optimisation
 * for its own sake: every character of ordinary Latin text reaches this on its
 * way to being measured, and the whole set begins above ASCII.
 */
export function isZeroAdvance(codePoint: number): boolean {
    if (codePoint < 0x00ad) {
        return false;
    }

    return DEFAULT_IGNORABLE.some(([first, last]) => codePoint >= first && codePoint <= last);
}

export class TrueTypeFont {
    private readonly hmtx: TableRecord;
    private readonly cmapSegments: CmapSegment[];
    private readonly glyphCache = new Map<number, number>();
    private kerningCache: GposKerning | null = null;

    private constructor(
        private readonly reader: FontReader,
        private readonly tables: Map<TableTag, TableRecord>,
        /** Font design units per em — the denominator for every advance. */
        readonly unitsPerEm: number,
        /** How many glyphs carry their own advance; the rest reuse the last. */
        private readonly numberOfHMetrics: number,
    ) {
        this.hmtx = requireTable(tables, 'hmtx');
        this.cmapSegments = readCmapFormat4(reader, tables);
    }

    static parse(bytes: Uint8Array): TrueTypeFont {
        const reader = new FontReader(bytes);
        const tables = readTableDirectory(reader);

        const head = requireTable(tables, 'head');
        const unitsPerEm = reader.uint16(head.offset + 18);
        if (unitsPerEm === 0) {
            throw new Error('Font declares unitsPerEm of 0; every advance would divide by zero.');
        }

        const hhea = requireTable(tables, 'hhea');
        const numberOfHMetrics = reader.uint16(hhea.offset + 34);
        if (numberOfHMetrics === 0) {
            throw new Error('Font declares no horizontal metrics; nothing can be measured with it.');
        }

        const font = new TrueTypeFont(reader, tables, unitsPerEm, numberOfHMetrics);
        // A font may ship neither table; the fallbacks put the rule a tenth of
        // the em below the baseline and half way up it, at a twentieth thick,
        // which is where a reader expects them and never zero — a rule of no
        // thickness is an underline that does not appear.
        const post = tables.get('post');
        const os2 = tables.get('OS/2');
        font.decoration = {
            underlinePositionUnits: undefined === post
                ? -Math.round(unitsPerEm / 10)
                : reader.int16(post.offset + 8),
            underlineThicknessUnits: undefined === post
                ? Math.round(unitsPerEm / 20)
                : reader.int16(post.offset + 10),
            strikeoutThicknessUnits: undefined === os2
                ? Math.round(unitsPerEm / 20)
                : reader.int16(os2.offset + 26),
            strikeoutPositionUnits: undefined === os2
                ? Math.round(unitsPerEm / 4)
                : reader.int16(os2.offset + 28),
        };

        const naming = tables.get('name');
        if (undefined !== naming) {
            // Name 16 is the TYPOGRAPHIC family and 1 the legacy one. They
            // differ exactly where it matters: Liberation Serif Bold calls
            // itself family "Liberation Serif" under 16 and "Liberation Serif
            // Bold" under 1, and the latter would have a renderer ask for a
            // family that does not exist and lose the weight as well.
            font.familyName = readName(reader, naming.offset, 16)
                ?? readName(reader, naming.offset, 1)
                ?? '';
            font.subfamilyName = readName(reader, naming.offset, 17)
                ?? readName(reader, naming.offset, 2)
                ?? '';
        }
        font.vertical = {
            // Descender is stored NEGATIVE (below the baseline) and is kept that
            // way; flipping its sign here would make every line height too
            // short by twice the descender, which is exactly the kind of error
            // that looks like "slightly wrong leading" rather than a bug.
            ascenderUnits: reader.int16(hhea.offset + 4),
            descenderUnits: reader.int16(hhea.offset + 6),
            lineGapUnits: reader.int16(hhea.offset + 8),
        };

        return font;
    }

    /**
     * Vertical metrics, in font units. The natural height of one line is
     * ascender − descender + lineGap, which is what "single spacing" means.
     */
    vertical: VerticalMetrics = { ascenderUnits: 0, descenderUnits: 0, lineGapUnits: 0 };

    /** Natural single-spaced line height at a size, in the same unit as the size. */
    naturalLineHeight(sizePx: number): number {
        const units = this.vertical.ascenderUnits - this.vertical.descenderUnits + this.vertical.lineGapUnits;

        return (units / this.unitsPerEm) * sizePx;
    }

    /** Distance from the line's top to the baseline, at a size. */
    ascent(sizePx: number): number {
        return (this.vertical.ascenderUnits / this.unitsPerEm) * sizePx;
    }

    /**
     * The family this file belongs to, as the FILE states it.
     *
     * Read from the font rather than taken from what the document asked for,
     * because those differ whenever a face was substituted: a renderer naming
     * the requested family would ask for a font that is not there, and get
     * metrics that are not the ones the layout measured with.
     *
     * Empty when the file carries no usable `name` table, which is legal and
     * leaves a renderer to fall back on a generic family.
     */
    familyName = '';

    /**
     * `Regular`, `Bold`, `Italic`, `Bold Italic` — how this file differs from
     * the family's regular face.
     *
     * Weight and slant live in SEPARATE FILES in this engine, so a renderer
     * cannot read them off a run; without this it would draw a bold heading in
     * the regular face while the layout measured the bold one.
     */
    subfamilyName = '';

    /**
     * Where a rule under the text goes, and how thick — the `post` table's
     * `underlinePosition` and `underlineThickness`, in font units.
     *
     * The position is NEGATIVE: it is below the baseline, and is kept that way
     * so a caller adds it rather than having to know which direction it means.
     *
     *  LibreOffice does not use these. Measured at 10pt it draws Liberation
     * Sans at −1.1/0.5 where the font says −0.33/0.73, Serif at −1.2/0.6
     * against −0.60/0.49, and Mono at −1.6/0.7 against −1.92/0.41 — the
     * thickness ordering is even inverted. Its source could not be identified
     * from three faces, so the FONT's own numbers are used, which is what the
     * designer specified and what a conformant renderer may use. The difference
     * is under half a point of position.
     *
     * A SECOND SIZE, and it still cannot be identified. Liberation
     * Serif at 40pt printed its underline at −4.40/2.20 and its strikeout at
     * +10.40/2.20, against this engine's −3.38/1.95 and +8.20/1.95. Neither
     * ratio holds across the two sizes — the underline's position is 1.43 times
     * ours at 10pt and 1.30 at 40 — so it is not a constant scale of the font's
     * numbers either, and the em fractions it looks like (0.11 and 0.055) do
     * not survive the 10pt row. The decision stands, now on six measurements
     * rather than three.
     */
    decoration: TextDecorationMetrics = {
        underlinePositionUnits: 0,
        underlineThicknessUnits: 0,
        strikeoutPositionUnits: 0,
        strikeoutThicknessUnits: 0,
    };

    /**
     * How far BELOW the baseline the centre of an underline sits, at a size.
     *
     * The table gives the TOP of the rule, not its middle — so half the
     * thickness is added, or a stroke drawn from this would sit half a rule too
     * high and a heavy underline would touch the text it belongs to.
     */
    underlineOffset(sizePx: number): number {
        const centre = this.decoration.underlinePositionUnits
            - this.decoration.underlineThicknessUnits / 2;

        return (-centre / this.unitsPerEm) * sizePx;
    }

    /** Where a rule THROUGH the text sits, above the baseline, at a size. */
    strikeoutOffset(sizePx: number): number {
        return (this.decoration.strikeoutPositionUnits / this.unitsPerEm) * sizePx;
    }

    underlineThickness(sizePx: number): number {
        return (this.decoration.underlineThicknessUnits / this.unitsPerEm) * sizePx;
    }

    strikeoutThickness(sizePx: number): number {
        return (this.decoration.strikeoutThicknessUnits / this.unitsPerEm) * sizePx;
    }

    /**
     * How far the font reaches BELOW the baseline, as a positive length.
     *
     * Positive because every caller wants a depth rather than a direction, and
     * `descenderUnits` is stored negative — a sign flipped at each call site is
     * a sign flipped wrongly at one of them.
     */
    descent(sizePx: number): number {
        return (-this.vertical.descenderUnits / this.unitsPerEm) * sizePx;
    }

    /** True when the file carries the legacy kerning table (Liberation does). */
    get hasLegacyKernTable(): boolean {
        return this.tables.has('kern');
    }

    /** True when kerning lives in GPOS, as it does for Carlito and Caladea. */
    get hasGposTable(): boolean {
        return this.tables.has('GPOS');
    }

    /**
     * Glyph id for a code point, or 0 (.notdef) when the font has no glyph.
     *
     * Zero is a real answer, not an error: a renderer draws .notdef and advances
     * by its width, so measurement must do the same or the line would be a
     * different length than the one on screen.
     */
    glyphId(codePoint: number): number {
        const cached = this.glyphCache.get(codePoint);
        if (cached !== undefined) {
            return cached;
        }

        let id = lookupFormat4(this.reader, this.cmapSegments, codePoint);
        // A character the font does not cover, but which another character IS:
        // see `SAME_SHAPE`. Resolved here rather than at reading time because
        // it is the FONT that is short of the glyph, and a font that has it
        // must draw its own.
        const instead = 0 === id ? SAME_SHAPE.get(codePoint) : undefined;
        if (undefined !== instead) {
            id = lookupFormat4(this.reader, this.cmapSegments, instead);
        }

        this.glyphCache.set(codePoint, id);

        return id;
    }

    /**
     * Advance width of a glyph in FONT UNITS.
     *
     * `hmtx` stores full metrics for the first `numberOfHMetrics` glyphs and
     * only left-side bearings after that — the trailing glyphs all share the
     * last recorded advance. Monospaced fonts exploit this heavily: Liberation
     * Mono records one advance and lets thousands of glyphs inherit it.
     */
    advanceWidthUnits(glyphId: number): number {
        const index = Math.min(glyphId, this.numberOfHMetrics - 1);
        const offset = this.hmtx.offset + index * 4;

        return this.reader.uint16(offset);
    }

    /**
     * Pair kerning for this font, parsed on first use.
     *
     * Lazy because parsing GPOS costs real work and a document that never
     * measures proportional text — or a caller that only wants raw advances —
     * should not pay for it.
     */
    get kerning(): GposKerning {
        if (null === this.kerningCache) {
            const gpos = this.tables.get('GPOS');
            this.kerningCache = gpos ? GposKerning.parse(this.reader, gpos) : GposKerning.none();
        }

        return this.kerningCache;
    }

    /**
     * Width of a string in POINTS at the given size, WITH pair kerning.
     *
     * NOT the number line breaking uses by itself. Which of this and
     * {@link measureUnkerned} applies is the DOCUMENT's to decide, through the
     * run's `w:kern`, and a run that says nothing is not kerned — measured off
     * a printed page, where `AVAVAVAVAV` came out 11.60pt wider without the
     * element than with it. Everything that measures a run goes
     * through `advanceOf`, which picks; call this one directly only when the
     * kerned width is what you actually want.
     *
     * Kerning is applied BETWEEN glyphs, so a single character never has any
     * and the adjustment is a property of the pair rather than of either glyph.
     * That is why this is not simply a sum and why {@link measureUnkerned}
     * still exists: the two differ, and the difference is the thing worth being
     * able to see.
     */
    measureAdvance(text: string, sizePt: number): MeasuredText {
        const glyphs: number[] = [];
        let units = 0;
        let missingGlyphs = 0;

        for (const character of text) {
            const codePoint = character.codePointAt(0);
            if (codePoint === undefined) {
                continue;
            }

            // Dropped rather than measured as zero, so the kern pair either
            // side of one still meets: `Ta` kerns the same with a bidi mark
            // between the letters as without. See {@link DEFAULT_IGNORABLE}.
            // It is also why a font that LACKS one reports no missing glyph —
            // `missingGlyphs` asks whether to substitute a face, and no face
            // would draw this character anyway.
            if (isZeroAdvance(codePoint)) {
                continue;
            }

            const glyph = this.glyphId(codePoint);
            if (glyph === 0) {
                missingGlyphs++;
            }
            glyphs.push(glyph);
            units += this.advanceWidthUnits(glyph);
        }

        const kerning = this.kerning;
        let kerningUnits = 0;
        if (!kerning.isEmpty) {
            for (let i = 1; i < glyphs.length; i++) {
                kerningUnits += kerning.between(glyphs[i - 1]!, glyphs[i]!);
            }
        }

        return {
            widthPt: ((units + kerningUnits) / this.unitsPerEm) * sizePt,
            units: units + kerningUnits,
            unkernedUnits: units,
            kerningUnits,
            missingGlyphs,
        };
    }

    /**
     * The same width with kerning deliberately switched off.
     *
     * The COMMON case, not a diagnostic: a run is kerned only where it says
     * `w:kern`, and nothing in the corpus does. It stayed a spare
     * primitive for a long time because the engine kerned everything — and it
     * is what proved the defect, since the printed page agreed with this
     * number and not with the kerned one.
     */
    measureUnkerned(text: string, sizePt: number): number {
        let units = 0;
        for (const character of text) {
            const codePoint = character.codePointAt(0);
            // The same skip as {@link measureAdvance}, and it has to be here
            // too: this is the COMMON path, since a run is kerned only where it
            // says `w:kern`. Fixing one and not the other would leave the
            // phantom width in place for almost every document.
            if (codePoint !== undefined && !isZeroAdvance(codePoint)) {
                units += this.advanceWidthUnits(this.glyphId(codePoint));
            }
        }

        return (units / this.unitsPerEm) * sizePt;
    }
}

/**
 * One string out of the `name` table, preferring the Windows/Unicode encoding.
 *
 * Records are UTF-16BE under platform 3 and single-byte under platform 1, and a
 * reader that assumed one of them gets either interleaved NULs or mojibake. The
 * whole table is scanned rather than indexed because records are sorted by
 * platform first, so the wanted id appears more than once.
 */
function readName(reader: FontReader, tableOffset: number, nameId: number): string | null {
    const count = reader.uint16(tableOffset + 2);
    const storage = tableOffset + reader.uint16(tableOffset + 4);
    let fallback: string | null = null;

    for (let index = 0; index < count; index++) {
        const record = tableOffset + 6 + index * 12;
        if (reader.uint16(record + 6) !== nameId) {
            continue;
        }

        const platform = reader.uint16(record);
        const length = reader.uint16(record + 8);
        const offset = storage + reader.uint16(record + 10);
        const wide = 3 === platform || 0 === platform;

        let text = '';
        for (let at = 0; at < length; at += wide ? 2 : 1) {
            text += String.fromCharCode(wide ? reader.uint16(offset + at) : reader.uint8(offset + at));
        }

        if (wide) {
            return text;
        }
        fallback ??= text;
    }

    return fallback;
}

/**
 * Where the rules of underlining and striking through go.
 *
 * In font units, so they scale with the size the way every other metric here
 * does. `underlinePositionUnits` is negative and `strikeoutPositionUnits`
 * positive, both as the tables store them.
 */
export interface TextDecorationMetrics {
    readonly underlinePositionUnits: number;
    readonly underlineThicknessUnits: number;
    readonly strikeoutPositionUnits: number;
    readonly strikeoutThicknessUnits: number;
}

export interface VerticalMetrics {
    readonly ascenderUnits: number;
    /** Negative — it is below the baseline, and stays that way. */
    readonly descenderUnits: number;
    readonly lineGapUnits: number;
}

export interface MeasuredText {
    /** Width in typographic points at the requested size, kerning included. */
    readonly widthPt: number;
    /** Total in font design units — size-independent, so safe to cache. */
    readonly units: number;
    /** The advance sum alone, before any pair adjustment. */
    readonly unkernedUnits: number;
    /**
     * What kerning contributed, almost always negative. Exposed because a
     * disagreement with another engine is usually a kerning disagreement, and
     * this says so immediately instead of requiring a second measurement.
     */
    readonly kerningUnits: number;
    /**
     * How many characters had no glyph in this font. Non-zero means the text
     * will render as .notdef boxes; the caller should substitute a font rather
     * than trust the width.
     */
    readonly missingGlyphs: number;
}

interface CmapSegment {
    readonly startCode: number;
    readonly endCode: number;
    readonly idDelta: number;
    readonly idRangeOffset: number;
    /** Address of this segment's idRangeOffset entry — the base for its trick. */
    readonly idRangeOffsetAddress: number;
}

/**
 * Read the cmap's format 4 (BMP) subtable.
 *
 * Format 4 is what every font we ship uses — verified against all twenty files
 * before this was written. A font carrying only format 12 (astral planes) would
 * fail loudly here rather than silently measure everything as .notdef.
 */
function readCmapFormat4(reader: FontReader, tables: Map<TableTag, TableRecord>): CmapSegment[] {
    const cmap = requireTable(tables, 'cmap');
    const subtableCount = reader.uint16(cmap.offset + 2);

    // Prefer Windows/BMP (3/1); fall back to Unicode (0/x). Both are format 4
    // in practice and point at the same data.
    let chosen = -1;
    for (let i = 0; i < subtableCount; i++) {
        const record = cmap.offset + 4 + i * 8;
        const platform = reader.uint16(record);
        const encoding = reader.uint16(record + 2);
        const offset = cmap.offset + reader.uint32(record + 4);

        if (reader.uint16(offset) !== 4) {
            continue;
        }
        if (platform === 3 && encoding === 1) {
            chosen = offset;
            break;
        }
        if (chosen < 0 && platform === 0) {
            chosen = offset;
        }
    }

    if (chosen < 0) {
        throw new Error('Font has no format 4 Unicode cmap subtable; its characters cannot be mapped to glyphs.');
    }

    const segCount = reader.uint16(chosen + 6) / 2;
    const endCodes = chosen + 14;
    const startCodes = endCodes + segCount * 2 + 2; // +2 skips reservedPad
    const idDeltas = startCodes + segCount * 2;
    const idRangeOffsets = idDeltas + segCount * 2;

    const segments: CmapSegment[] = [];
    for (let i = 0; i < segCount; i++) {
        segments.push({
            endCode: reader.uint16(endCodes + i * 2),
            startCode: reader.uint16(startCodes + i * 2),
            idDelta: reader.int16(idDeltas + i * 2),
            idRangeOffset: reader.uint16(idRangeOffsets + i * 2),
            idRangeOffsetAddress: idRangeOffsets + i * 2,
        });
    }

    return segments;
}

/**
 * Format 4 lookup.
 *
 * The `idRangeOffset` branch is the awkward part of the format: a non-zero
 * value is a byte offset measured FROM THE OFFSET'S OWN ADDRESS into a glyph
 * array that follows the segment arrays. Computing it from anywhere else gives
 * plausible-looking wrong glyphs, which is why the segment carries its own
 * address rather than an index.
 */
function lookupFormat4(reader: FontReader, segments: CmapSegment[], codePoint: number): number {
    if (codePoint > 0xffff) {
        // Beyond the BMP: format 4 cannot express it, so the font has no glyph.
        return 0;
    }

    for (const segment of segments) {
        if (segment.endCode < codePoint) {
            continue;
        }
        if (segment.startCode > codePoint) {
            return 0;
        }

        if (segment.idRangeOffset === 0) {
            return (codePoint + segment.idDelta) & 0xffff;
        }

        const address = segment.idRangeOffsetAddress
            + segment.idRangeOffset
            + (codePoint - segment.startCode) * 2;
        const glyph = reader.uint16(address);

        return glyph === 0 ? 0 : (glyph + segment.idDelta) & 0xffff;
    }

    return 0;
}
