import { XmlDocument, type XmlElement } from '../ooxml/xml.js';
import {
    eighthPointsToPx,
    emuToPx,
    halfPointsToPx,
    pointsToPx,
    pointsToTwips,
    pxToPoints,
    twipsToPx,
} from '../ooxml/units.js';
import type { BorderSide, BorderStyle, BoxBorders } from '../layout/borders.js';
import { formatNumeral, type NumeralStyle } from '../text/numerals.js';
import type {
    BoxInset,
    Float,
    FloatingBox,
    FloatingImage,
    FloatPosition,
    RelativeTo,
    WrapMode,
} from '../layout/float.js';
import {
    OBJECT_REPLACEMENT,
    type ImageContent,
    type InlineImage,
    type InlineShape,
} from '../layout/image.js';
import { advanceOf, type FieldKind, type RubyGloss, type StyledRun }
    from '../layout/line-breaker.js';
import { SOFT_HYPHEN } from '../layout/break-opportunities.js';
import { DEFAULT_TAB_PX, isTable, stackBlocks } from '../layout/page-layout.js';
import type {
    Block,
    ContentBox,
    FurnitureSource,
    ListMarker,
    Section,
    PageColumn,
    PageGeometry,
    Paragraph,
    ParagraphStyle,
    Table,
    CellMargins,
    CellTextDirection,
    CellVerticalAlign,
    LineRule,
    LineNumbering,
    PageBorders,
    TableAlignment,
    TableCell,
    TableRow,
    VerticalMerge,
} from '../layout/page-layout.js';
import type { FontCatalogue } from './font-catalogue.js';
import {
    StyleSheet,
    readParagraphProperties,
    readRunProperties,
    type ParagraphProperties,
    type RunProperties,
    readBorderSide,
    type ConditionalFormat,
} from './style-sheet.js';
import { translateLine, translateRow } from '../layout/table-layout.js';
import { Numbering, NumberingCounters, type NumberingLevel } from './numbering.js';

/**
 * Reading `word/document.xml` into the paragraphs the layout engine flows.
 *
 * This is where a real document first meets the engine: the XML layer supplies
 * the tree, the style sheet says what each paragraph's formatting resolves to,
 * the font catalogue turns "Calibri" into bytes we hold, and the result is the
 * exact `Paragraph[]` and `PageGeometry` that `layoutPages` already takes.
 */

export type DiagnosticKind =
    | 'font-substituted'
    | 'revision-hidden'
    | 'unknown-numbering'
    | 'unknown-paragraph-style'
    | 'unsupported-block'
    | 'unsupported-section-break';

/** A run of blocks sharing one page setup. */
export interface DocumentSection extends Omit<Section, 'blocks'> {
    readonly blocks: Block[];
    /** Index of this section's first block in the document's flat list. */
    readonly firstBlockIndex: number;
    /** This section's own header and footer, which need not be the document's. */
    readonly headers: ReadonlyMap<FurnitureVariant, PageFurniture>;
    readonly footers: ReadonlyMap<FurnitureVariant, PageFurniture>;
    readonly variantForPage: (pageIndexInSection: number, pageNumber: number) => FurnitureVariant;
}

export interface Diagnostic {
    readonly kind: DiagnosticKind;
    readonly detail: string;
}

export interface WordDocument {
    /**
     * The document's blocks: paragraphs and tables, in reading order.
     *
     * Named `paragraphs` before it could hold tables, and kept that way — it
     * is what `layoutPages` takes, and the type says what is in it.
     */
    readonly paragraphs: Block[];
    readonly geometry: PageGeometry;
    /**
     * What could not be honoured. Reported rather than swallowed: a document
     * that silently lost its tables still lays out, and looks correct, and is
     * wrong — and the page count is the only clue.
     */
    readonly diagnostics: Diagnostic[];
    readonly headers: ReadonlyMap<FurnitureVariant, PageFurniture>;
    readonly footers: ReadonlyMap<FurnitureVariant, PageFurniture>;
    /**
     * Which variant a page draws, given its index within its section AND its
     * printed page number. Both are needed; see the implementation for why.
     */
    readonly variantForPage: (pageIndexInSection: number, pageNumber: number) => FurnitureVariant;
    /**
     * The writing area of page N, ready to hand to `layoutPages`.
     *
     * A header only shrinks the page when it is TALLER than the margin it
     * sits in: Word draws it in the margin and pushes the body down only
     * once it runs out of room there.
     */
    readonly contentBox: (pageIndex: number, pageNumber: number) => ContentBox;
    /**
     * The document's sections, each with its own page setup.
     *
     * Always at least one. Hand these to `layoutSections` to honour a
     * landscape page in the middle of a portrait document; `paragraphs`
     * remains the flat list of every block, in reading order.
     */
    readonly sections: DocumentSection[];
    /**
     * Each footnote's own blocks, by the `w:id` its references name.
     *
     * Separately from the body because a note is not IN the flow: it belongs to
     * whichever page its reference lands on, which nothing knows until the
     * text has been broken into pages.
     */
    readonly footnotes: ReadonlyMap<number, Block[]>;
}

/** Which header or footer a page uses. */
export type FurnitureVariant = 'default' | 'first' | 'even';

export interface PageFurniture {
    /** The header's or footer's own blocks, in reading order. */
    readonly blocks: Block[];
    /** How tall it came out at the writing width. */
    readonly heightPx: number;
}

export interface FurnitureParts {
    /** Variant to the part's XML. A document may declare any subset. */
    readonly headers?: Partial<Record<FurnitureVariant, string>>;
    readonly footers?: Partial<Record<FurnitureVariant, string>>;
    /**
     * Every header and footer part in the package, by relationship id.
     *
     * Supplied this way because only the READER knows which section uses
     * which: a `w:headerReference` names an `r:id`, and the references live
     * in the section properties the reader is already walking. The caller
     * would have to split the body a second time to work it out, and two
     * implementations of that would eventually disagree.
     */
    readonly furnitureById?: Readonly<Record<string, string>>;
}

export interface ReadOptions extends FurnitureParts {
    readonly documentXml: string;
    /** `styles.xml`. Absent is legal; the document then gets bare defaults. */
    readonly stylesXml?: string;
    /** `numbering.xml`. Absent is legal; the document then has no lists. */
    readonly numberingXml?: string;
    /** `settings.xml`, which is where a document says it uses even-page headers. */
    readonly settingsXml?: string;
    /** `footnotes.xml`. Absent is legal; the document then has no notes. */
    readonly footnotesXml?: string;
    /** `endnotes.xml`. Absent is legal; the document then has none. */
    readonly endnotesXml?: string;
    /**
     * Every picture in the package, by relationship id.
     *
     * Supplied the same way the headers are, and for the same reason: only the
     * reader knows which `r:embed` belongs to which run, and it is already
     * walking the drawing that names it.
     */
    readonly mediaById?: Readonly<Record<string, ImageContent>>;
    readonly fonts: FontCatalogue;
}

/** A4 with one-inch margins — what a section that declares nothing gets. */
const DEFAULT_GEOMETRY = {
    widthTwips: 11906,
    heightTwips: 16838,
    marginTwips: 1440,
} as const;

export function readWordDocument(options: ReadOptions): WordDocument {
    const root = XmlDocument.parse(options.documentXml).root;
    const body = root.element('w:body');
    if (null === body) {
        throw new Error('word/document.xml has no <w:body>; it is not a WordprocessingML document.');
    }

    const styles = undefined === options.stylesXml
        ? StyleSheet.empty()
        : StyleSheet.parse(options.stylesXml);

    const numbering = undefined === options.numberingXml
        ? Numbering.empty()
        : Numbering.parse(options.numberingXml);

    // `w:defaultTabStop` is document-wide and is what every tab past the last
    // explicit stop advances by.
    const settings = undefined === options.settingsXml
        ? null
        : XmlDocument.parse(options.settingsXml).root;
    const defaultTabTwips = Number(
        settings?.element('w:defaultTabStop')?.attribute('w:val') ?? '',
    );
    const defaultTabPx = Number.isFinite(defaultTabTwips) && defaultTabTwips > 0
        ? twipsToPx(defaultTabTwips)
        : DEFAULT_TAB_PX;

    // `w:decimalSymbol` is document-wide, and it is a COMMA across most of
    // Europe: a document that says so and is aligned on a full stop has its
    // numbers ranged on nothing at all.
    const decimalSymbol = settings?.element('w:decimalSymbol')?.attribute('w:val') ?? null;

    const reader = new BodyReader(
        styles,
        numbering,
        options.fonts,
        defaultTabPx,
        options.mediaById ?? {},
        decimalSymbol,
    );

    const raw = reader.readSections(body);

    // AFTER the body: a note's number is its place among the body's references,
    // so nothing about the notes can be settled until the body has been walked.
    const footnotes = new Map<number, Block[]>();
    if (undefined !== options.footnotesXml) {
        const byId = new Map<number, XmlElement>();
        // Every `w:footnote` in the part, including the `separator` and
        // `continuationSeparator` entries — which no body reference names, so
        // the loop below never asks for them. They hold the rule Word draws
        // above the block, which this engine draws itself.
        //
        // MEASURED, where this was a claim: the same document with
        // its separator paragraph emptied, and again with both separator
        // footnotes removed outright, printed the identical rule — 28.350 to
        // 172.350 at 0.100. The rule is the RENDERER's furniture and a file
        // cannot change it, so leaving these two unread is right rather than
        // merely convenient.
        for (const part of XmlDocument.parse(options.footnotesXml).root.elements('w:footnote')) {
            const id = attributeNumber(part, 'w:id');
            if (null !== id) {
                byId.set(id, part);
            }
        }

        for (const id of reader.referencedFootnotes()) {
            const part = byId.get(id);
            if (undefined === part) {
                reader.report('unsupported-block', `w:footnoteReference "${id}" has no footnote`);
                continue;
            }
            footnotes.set(id, reader.readFootnote(id, part));
        }
    }

    // Endnotes are not footnotes: they belong to the END of the document
    // rather than to the page their reference fell on, so they are simply the
    // last blocks in it and need nothing of the page flow at all.
    const endnoteBlocks: Block[] = [];
    if (undefined !== options.endnotesXml) {
        const byId = new Map<number, XmlElement>();
        for (const part of XmlDocument.parse(options.endnotesXml).root.elements('w:endnote')) {
            const id = attributeNumber(part, 'w:id');
            if (null !== id) {
                byId.set(id, part);
            }
        }

        for (const id of reader.referencedEndnotes()) {
            const part = byId.get(id);
            if (undefined === part) {
                reader.report('unsupported-block', `w:endnoteReference "${id}" has no endnote`);
                continue;
            }
            endnoteBlocks.push(...reader.readEndnote(id, part));
        }
    }

    const lastPart = raw[raw.length - 1];
    if (undefined !== lastPart && endnoteBlocks.length > 0) {
        lastPart.blocks.push(...endnoteBlocks);
    }
    const sectPr = body.element('w:sectPr');
    const geometry = readGeometry(sectPr);
    const writingWidth = geometry.widthPx - geometry.marginLeftPx - geometry.marginRightPx;

    const measure = (xml: string): PageFurniture => {
        const blocks = reader.read(XmlDocument.parse(xml).root);

        return { blocks, heightPx: stackBlocks(blocks, writingWidth).heightPx };
    };

    const byVariant = (parts: Partial<Record<FurnitureVariant, string>> | undefined)
        : Map<FurnitureVariant, PageFurniture> => {
        const out = new Map<FurnitureVariant, PageFurniture>();
        for (const [variant, xml] of Object.entries(parts ?? {})) {
            out.set(variant as FurnitureVariant, measure(xml));
        }

        return out;
    };

    /**
     * Re-measure furniture at the writing width of the section that draws it.
     *
     * A header's height is what pushes the body down, and it depends on the
     * column the header wraps in. A landscape section inheriting a portrait
     * section's header must not inherit its height along with it: the same
     * words occupy fewer lines on wider paper, and the body would start too far
     * down for the rest of the document.
     *
     * Only the height is recomputed — the blocks were parsed once and do not
     * change with the paper.
     */
    const remeasured = (
        map: ReadonlyMap<FurnitureVariant, PageFurniture>,
        widthPx: number,
    ): Map<FurnitureVariant, PageFurniture> => {
        const out = new Map<FurnitureVariant, PageFurniture>();
        for (const [variant, furniture] of map) {
            out.set(variant, {
                blocks: furniture.blocks,
                heightPx: stackBlocks(furniture.blocks, widthPx).heightPx,
            });
        }

        return out;
    };

    /**
     * A section's own references, resolved through the parts the caller
     * supplied. A variant the section does not name is INHERITED from the
     * section before it, which is what Word does — only the first section of
     * a document has to declare everything it uses.
     */
    const referenced = (
        element: string,
        own: XmlElement | null,
        inherited: ReadonlyMap<FurnitureVariant, PageFurniture>,
    ): Map<FurnitureVariant, PageFurniture> => {
        const out = new Map(inherited);
        for (const reference of own?.elements(element) ?? []) {
            const id = reference.attribute('r:id');
            const xml = null === id ? undefined : options.furnitureById?.[id];
            if (undefined === xml) {
                continue;
            }
            out.set((reference.attribute('w:type') ?? 'default') as FurnitureVariant, measure(xml));
        }

        return out;
    };

    /**
     * `w:evenAndOddHeaders` — OFF by default, and it is a document-wide
     * setting rather than a section one.
     *
     * Without it Word ignores even-page headers ALTOGETHER, even when a
     * section defines one. Using them regardless would give every other page
     * a header the document asked not to show.
     */
    const evenAndOdd = null !== (settings?.element('w:evenAndOddHeaders') ?? null);

    /**
     * Which header or footer a page draws.
     *
     * The two arguments are not interchangeable and both are needed.
     * `w:titlePg` picks out the first page of THIS SECTION, while
     * `w:evenAndOddHeaders` alternates on the PRINTED page number — so a
     * section opening on document page four is its own first page and an even
     * page at the same time.
     *
     * Verified against LibreOffice on exactly that page: with `w:titlePg` it
     * draws the new section's FIRST header, and without it the new section's
     * EVEN header. Counting parity from the section instead would give it the
     * default one, which is neither.
     */
    const variantFor = (own: XmlElement | null) =>
        (pageIndexInSection: number, pageNumber: number): FurnitureVariant => {
            if (null !== (own?.element('w:titlePg') ?? null) && 0 === pageIndexInSection) {
                return 'first';
            }

            return evenAndOdd && 0 === pageNumber % 2 ? 'even' : 'default';
        };

    /**
     * The furniture a page actually draws, falling back to the default variant
     * when the one it asked for was never declared.
     *
     * One function rather than a rule repeated in two places: the height the
     * body is pushed down by and the blocks that get drawn have to come from
     * the SAME header, and two copies of the fallback would eventually pick
     * different ones.
     */
    const furnitureAt = (
        map: ReadonlyMap<FurnitureVariant, PageFurniture>,
        variant: (pageIndexInSection: number, pageNumber: number) => FurnitureVariant,
    ) => (pageIndexInSection: number, pageNumber: number): PageFurniture | undefined =>
        map.get(variant(pageIndexInSection, pageNumber)) ?? map.get('default');

    const blocksFor = (
        map: ReadonlyMap<FurnitureVariant, PageFurniture>,
        variant: (pageIndexInSection: number, pageNumber: number) => FurnitureVariant,
    ): FurnitureSource => {
        const at = furnitureAt(map, variant);

        return (pageIndexInSection, pageNumber) => at(pageIndexInSection, pageNumber)?.blocks;
    };

    const boxFor = (
        own: PageGeometry,
        headerMap: ReadonlyMap<FurnitureVariant, PageFurniture>,
        footerMap: ReadonlyMap<FurnitureVariant, PageFurniture>,
        variant: (pageIndexInSection: number, pageNumber: number) => FurnitureVariant,
    ) => {
        const headerAt = furnitureAt(headerMap, variant);
        const footerAt = furnitureAt(footerMap, variant);

        return (pageIndex: number, pageNumber: number): ContentBox => {
            const headerBottom = (own.headerDistancePx ?? 0)
                + (headerAt(pageIndex, pageNumber)?.heightPx ?? 0);
            const footerTop = own.heightPx - (own.footerDistancePx ?? 0)
                - (footerAt(pageIndex, pageNumber)?.heightPx ?? 0);

            return {
                topPx: Math.max(own.marginTopPx, headerBottom),
                bottomPx: Math.min(own.heightPx - own.marginBottomPx, footerTop),
            };
        };
    };

    // A `continuous` break does not start a page, so its blocks belong to the
    // section before it. Folding them keeps the behaviour a document without
    // sections already had, and only the paper it wanted to change is lost —
    // which is reported rather than quietly applied a page too early.
    const sections: DocumentSection[] = [];
    let firstBlockIndex = 0;
    let inheritedHeaders: ReadonlyMap<FurnitureVariant, PageFurniture> = byVariant(options.headers);
    let inheritedFooters: ReadonlyMap<FurnitureVariant, PageFurniture> = byVariant(options.footers);

    for (const part of raw) {
        const partGeometry = readGeometry(part.sectPr ?? sectPr);
        const continuous = 'continuous' === (part.sectPr?.element('w:type')?.attribute('w:val') ?? null);
        const previous = sections[sections.length - 1];

        if (continuous && undefined !== previous) {
            if (!sameGeometry(previous.geometry, partGeometry)) {
                reader.report(
                    'unsupported-section-break',
                    'a continuous section break changed the page setup; the previous one was kept',
                );
            }
            previous.blocks.push(...part.blocks);

            continue;
        }

        const own = part.sectPr;
        const partWidth = partGeometry.widthPx - partGeometry.marginLeftPx - partGeometry.marginRightPx;
        const partHeaders = remeasured(referenced('w:headerReference', own, inheritedHeaders), partWidth);
        const partFooters = remeasured(referenced('w:footerReference', own, inheritedFooters), partWidth);
        const partVariant = variantFor(own);
        inheritedHeaders = partHeaders;
        inheritedFooters = partFooters;

        const declared = part.sectPr?.element('w:type')?.attribute('w:val') ?? null;
        const restart = readFirstPageNumber(part.sectPr);
        const numberFormat = readPageNumberFormat(part.sectPr);
        const pageBorders = readPageBorders(part.sectPr);
        const lineNumbering = readLineNumbering(part.sectPr);
        const pitchPx = readLinePitchPx(part.sectPr);

        sections.push({
            blocks: null === pitchPx
                ? part.blocks
                : onGrid(
                    part.blocks,
                    pitchPx,
                    (block) => reader.snapsToGrid(block),
                    // `null`, not `undefined` — the sentinel above. Written
                    // the other way this read TRUE for every document and put
                    // them all on the modern branch, which the no-settings
                    // fixture caught at once.
                    null !== settings,
                ),
            geometry: partGeometry,
            ...(null === restart ? {} : { firstPageNumber: restart }),
            ...(null === numberFormat ? {} : { pageNumberFormat: numberFormat }),
            ...(null === pageBorders ? {} : { pageBorders }),
            ...(null === lineNumbering ? {} : { lineNumbering }),
            contentBox: boxFor(partGeometry, partHeaders, partFooters, partVariant),
            // Anything else — nextColumn, or nothing at all — starts a page,
            // which is what a section break does when it says no more.
            startsOn: 'evenPage' === declared || 'oddPage' === declared ? declared : 'nextPage',
            firstBlockIndex,
            headers: partHeaders,
            footers: partFooters,
            variantForPage: partVariant,
            headerFor: blocksFor(partHeaders, partVariant),
            footerFor: blocksFor(partFooters, partVariant),
        });
        firstBlockIndex += part.blocks.length;
    }

    // The document-level view is the FIRST section's: it is the one a caller
    // that knows nothing about sections is looking at.
    const first = sections[0];
    const headers = first?.headers ?? new Map<FurnitureVariant, PageFurniture>();
    const footers = first?.footers ?? new Map<FurnitureVariant, PageFurniture>();
    const variantForPage = first?.variantForPage ?? ((): FurnitureVariant => 'default');
    const contentBox = first?.contentBox ?? boxFor(geometry, headers, footers, variantForPage);

    return {
        paragraphs: sections.flatMap((section) => section.blocks),
        geometry,
        diagnostics: reader.diagnostics,
        headers,
        footers,
        variantForPage,
        contentBox,
        sections: sections.map((section) => ({ ...section, footnotes })),
        footnotes,
    };
}

/**
 * How much smaller a small capital is than a full one.
 *
 * Measured against LibreOffice: six lower-case letters set in small caps came
 * to the width of the same six capitals at four fifths of the size, and a
 * capital already in the run was untouched.
 */
/**
 * `w:sym/@w:char` — a hexadecimal code, and where in Unicode it lands.
 *
 * Word writes symbol characters in the F000 private-use block: the code is the
 * font's own byte with F000 added, which is a convention rather than a Unicode
 * meaning. Taken back out, so the character asked for is the one looked up in
 * the font the document named.
 */
function symbolCodePoint(char: string | null): number | null {
    if (null === char) {
        return null;
    }

    const point = Number.parseInt(char, 16);
    if (!Number.isFinite(point) || point <= 0) {
        return null;
    }

    return point >= 0xF000 && point <= 0xF0FF ? point - 0xF000 : point;
}

/**
 * What an endnote's mark counts in.
 *
 * Word's default for endnotes, and what LibreOffice printed: lower-case roman,
 * which is what tells a reader that an `i` in the text is an endnote where a
 * `1` is a footnote.
 */
const ENDNOTE_NUMERALS: NumeralStyle = 'lowerRoman';

export const SMALL_CAPS_SCALE = 0.8;

/**
 * A small caps run, split where its case changes.
 *
 * One run cannot be two sizes, so a run that mixes cases has to become
 * several: `Abc` is a full-size `A` and a reduced `BC`. Splitting is what makes
 * the WIDTH right, which is why it happens here rather than at drawing time.
 *
 * Case is decided by whether uppercasing the character changes it — which
 * leaves digits, spaces and punctuation at full size, where they belong.
 */
function smallCapped(text: string): { text: string; small: boolean }[] {
    const pieces: { text: string; small: boolean }[] = [];

    for (const character of text) {
        const small = character !== character.toUpperCase();
        const last = pieces[pieces.length - 1];
        if (undefined !== last && last.small === small) {
            last.text += character;
        } else {
            pieces.push({ text: character, small });
        }
    }

    return pieces;
}

/**
 * The `\*` switch on a field instruction, if it names a numbering.
 *
 * The same instruction can carry several switches, and most of them are not
 * about numbering at all — `\* MERGEFORMAT` says to keep the field's
 * formatting when it updates, and treating it as a numbering would print a
 * digit where the document asked for one anyway but by accident.
 *
 * Verified against LibreOffice, which prints `iii` for `\* roman`, `III` for
 * `\* ROMAN`, `c` for `\* alphabetic`, `C` for `\* ALPHABETIC` and a plain
 * `3` for `\* MERGEFORMAT`.
 */
function fieldFormat(instruction: string | null): NumeralStyle | null {
    for (const match of (instruction ?? '').matchAll(/\\\*\s*(\w+)/gu)) {
        const style = FIELD_NUMERALS[match[1] ?? ''];
        if (undefined !== style) {
            return style;
        }
    }

    return null;
}

const FIELD_NUMERALS: Record<string, NumeralStyle | undefined> = {
    Arabic: 'decimal',
    arabic: 'decimal',
    ARABIC: 'decimal',
    roman: 'lowerRoman',
    Roman: 'lowerRoman',
    ROMAN: 'upperRoman',
    alphabetic: 'lowerLetter',
    Alphabetic: 'lowerLetter',
    ALPHABETIC: 'upperLetter',
};

/**
 * How much smaller a superscript or subscript run is, and how far it moves.
 *
 * Measured out of LibreOffice at ten point: it draws both scripts at 5.8pt,
 * raises a superscript 3.95pt and drops a subscript 0.9pt. Word's own defaults
 * differ — it documents 65% and a third — so a document round-tripped through
 * both will not agree with either exactly.
 */
const SCRIPT_SCALE = 0.58;
const SUPERSCRIPT_RISE = 0.395;
const SUBSCRIPT_DROP = 0.09;

/**
 * A list marker that has been rendered but not yet PLACED.
 *
 * Its position needs the paragraph's merged indents and tab stops, which are
 * not known while the numbering is read — see {@link WordDocumentReader.placeMarker}.
 */
interface PendingMarker {
    readonly run: StyledRun;
    /** What is DRAWN, so it is what a right or centred marker is justified by. */
    readonly widthPx: number;
    /** The width a `space` suffix adds after the marker; zero for the others. */
    readonly spacePx: number;
    readonly justification: 'left' | 'right' | 'center';
    readonly suffix: 'tab' | 'space' | 'nothing';
}

/**
 * A level's `w:ind`, expressed exactly as a paragraph's own `w:ind` would be.
 *
 * A hanging indent is a negative first-line indent in both places, so offering
 * it in that form lets the ordinary style/numbering/direct merge pick a winner
 * — rather than the reader deciding, which is how a paragraph's own indent came
 * to be applied ON TOP of the level's.
 */
function levelIndents(level: NumberingLevel): ParagraphProperties {
    return {
        ...pick('indentLeftTwips', level.indentLeftTwips),
        ...pick('indentFirstLineTwips', null === level.indentHangingTwips
            ? null
            : -level.indentHangingTwips),
    };
}

/** How far through a complex field the run walker currently is. */
interface FieldState {
    /** The field being resulted, once its instruction has been read. */
    kind: FieldKind | null;
    /** How that field writes its number, from the instruction's `\*` switch. */
    format: NumeralStyle | null;
    /** The instruction so far, or null when not inside a field at all. */
    instruction: string | null;
    /** True between `separate` and `end`, where the cached value lives. */
    inResult: boolean;
}

/**
 * Which field an instruction asks for, if it is one this engine answers.
 *
 * The instruction is free text — ` PAGE `, ` NUMPAGES  \* MERGEFORMAT ` — so
 * the FIRST word names the field and the rest are switches. Anything else, a
 * `TOC` or a `REF` or a date, returns null and keeps the value Word cached:
 * the best available answer for a field this engine cannot compute.
 */
function fieldKind(instruction: string | null): FieldKind | null {
    const name = (instruction ?? '').trim().split(/\s+/)[0]?.toUpperCase() ?? '';

    if ('PAGE' === name) {
        return 'page';
    }

    return 'NUMPAGES' === name ? 'numPages' : null;
}

/**
 * One `w:*Borders` element.
 *
 * Returns undefined when nothing is declared, so an absent element and one full
 * of `nil` sides stay distinguishable from a table that asked for no borders at
 * all — both draw nothing, but only the second overrides an inherited style.
 */
function readBorders(element: XmlElement | null): BoxBorders | undefined {
    return readBordersFrom(null === element ? [] : [element]);
}

/**
 * The same, merged across a chain of border elements — a table style's, then
 * the ones it is based on, then the table's own.
 *
 * ## Per SIDE, not per element
 *
 * Measured against LibreOffice: a table naming a style that draws
 * 1.5pt all round, and declaring nothing but its own `insideH` at 3pt, printed
 * the middle rule at 3.000 and the four outer ones at 1.500. So the table's
 * element does not replace the style's — each side is answered by the LAST
 * level that names it.
 *
 * A side named as `none` is an answer too, and turns the side off: the same
 * table declaring all six that way printed nothing at all, where the style
 * alone would have drawn a grid. That is why the merge walks ELEMENTS rather
 * than the parsed sides, which cannot tell "off" from "unmentioned".
 */
function readBordersFrom(
    elements: readonly (XmlElement | null)[],
): BoxBorders | undefined {
    const present = elements.filter((element): element is XmlElement => null !== element);
    if (0 === present.length) {
        return undefined;
    }

    const sides: Record<string, BorderSide> = {};
    for (const [name, tag] of Object.entries(BORDER_SIDES)) {
        for (const element of present) {
            const declared = element.element(tag) ?? null;
            if (null === declared) {
                continue;
            }

            const side = readBorderSide(declared);
            if (null === side) {
                delete sides[name];
            } else {
                sides[name] = side;
            }
        }
    }

    // A paragraph's `w:between` is the rule it shares with the paragraph below
    // it, which is the same place a table's inside-horizontal rule sits — so
    // it is read into that side rather than given one of its own.
    for (const element of present) {
        const between = readBorderSide(element.element('w:between') ?? null);
        if (null !== between && undefined === sides['insideH']) {
            sides['insideH'] = between;
        }
    }

    return 0 === Object.keys(sides).length ? undefined : sides;
}

const BORDER_SIDES: Record<string, string> = {
    top: 'w:top',
    left: 'w:left',
    bottom: 'w:bottom',
    right: 'w:right',
    insideH: 'w:insideH',
    insideV: 'w:insideV',
};

/**
 * OOXML's thirty-odd border styles, folded onto the four this engine draws.
 *
 * Anything unrecognised becomes `solid`: an unknown style is still a LINE, and
 * drawing nothing would lose a border the document asked for — the width and
 * the presence are what move the eye, not whether the flourish is right.
 */
/**
 * `w:shd/@w:fill` — the colour BEHIND the text.
 *
 * `w:color` on the same element is the pattern's foreground, which only shows
 * for the hatched `w:val` patterns this engine does not draw; taking it would
 * fill a cell with the wrong colour entirely.
 */
/**
 * `w:vAlign` — where a cell's content sits in a row taller than it.
 *
 * `both` means vertically JUSTIFIED, spreading the paragraphs out to fill the
 * cell. That is a different operation from moving them as a block, so it is
 * read as the top alignment it degrades to rather than mistaken for `bottom`.
 */
function readCellVerticalAlign(element: XmlElement | null): CellVerticalAlign | null {
    const value = element?.attribute('w:val') ?? null;

    return 'center' === value || 'bottom' === value || 'top' === value ? value : null;
}

/**
 * `w:tcMar` — the sides this cell pads itself by, and only those.
 *
 * A side left out is not nought: it keeps the table's `w:tblCellMar`, which is
 * why this returns the sides it found rather than four numbers. Measured
 * against LibreOffice, an override on one cell left its neighbour alone.
 */
function readCellMargins(element: XmlElement | null): CellMargins | undefined {
    if (null === element) {
        return undefined;
    }

    const sides = ['left', 'right', 'top', 'bottom'] as const;
    const margins: Record<string, number> = {};

    for (const side of sides) {
        const declared = element.element(`w:${side}`);
        const value = attributeNumber(declared, 'w:w');

        // `w:type` may say `nil`, which is nought rather than absent — the
        // cell is stating that it pads itself by nothing at all.
        if (null === declared || null === value) {
            continue;
        }
        margins[`${side}Px`] = 'nil' === declared.attribute('w:type')
            ? 0
            : twipsToPx(value);
    }

    return 0 === Object.keys(margins).length ? undefined : margins;
}

/**
 * `w:tblW`/`w:tcW` percentages count FIFTIETHS of a percent, so 5000 is all
 * of it — defined by the OOXML schema, not measured off a page.
 */
const PERCENT_UNITS = 5000;

/**
 * `w:tblW` — what the table asks to be, in whichever unit it asked.
 *
 * A fraction cannot be turned into a width here: the reader does not know what
 * the table will land in, and the same table is one width on the page and
 * another inside a cell. Measured against LibreOffice, a fifty-percent table
 * came out at 269.29pt on an A4 text column of 538.58 — half of it, exactly.
 */
function readPreferredWidth(
    element: XmlElement | null,
): { twips: number | null; fraction: number | null } {
    const value = attributeNumber(element, 'w:w');
    const type = element?.attribute('w:type') ?? 'dxa';

    if (null === element || null === value || value <= 0 || 'auto' === type) {
        return { twips: null, fraction: null };
    }

    return 'pct' === type
        ? { twips: null, fraction: value / PERCENT_UNITS }
        : { twips: value, fraction: null };
}

/**
 * `w:tblPr/w:jc` — where the table sits in the column it lands in.
 *
 * `start` and `end` are the bidirectional spellings of `left` and `right`; in
 * a left-to-right document they mean the same thing, which is the only
 * direction this lays out.
 */
function readTableAlignment(element: XmlElement | null): TableAlignment | null {
    const value = element?.attribute('w:val') ?? null;

    if ('center' === value) {
        return 'center';
    }
    if ('right' === value || 'end' === value) {
        return 'right';
    }

    return 'left' === value || 'start' === value ? 'left' : null;
}

/**
 * A VML colour, as `#RRGGBB`.
 *
 * VML writes `#RRGGBB`, and also the short `#RGB` and bare colour NAMES that
 * CSS allows. Only the two hexadecimal spellings are understood; a name is
 * refused rather than guessed at, which leaves the shape unfilled instead of
 * filled in a colour nobody asked for.
 */
function normaliseVmlColour(colour: string): string | null {
    const value = colour.trim();
    if (/^#[0-9a-f]{6}$/iu.test(value)) {
        return value.toUpperCase();
    }
    if (/^#[0-9a-f]{3}$/iu.test(value)) {
        const [, r, g, b] = value.toUpperCase();

        return `#${r!}${r!}${g!}${g!}${b!}${b!}`;
    }

    return null;
}

/**
 * The block-level children of an element, with content controls flattened.
 *
 * `w:sdt` is a WRAPPER and nothing more to a reader: a template's fillable
 * region, a repeating section, a date picker. What it holds are ordinary
 * paragraphs, tables and rows, and they belong exactly where the control sits.
 * A walker that knew only `w:p` and `w:tbl` dropped every one of them — which
 * in a template is most of the document, since a control is how a template
 * says where its content goes.
 *
 * Flattened recursively, because a repeating section is a control whose rows
 * are controls.
 */
function* blockChildren(parent: XmlElement): Generator<XmlElement> {
    for (const child of parent.elements()) {
        if ('w:sdt' === child.name) {
            const content = child.element('w:sdtContent');
            if (null !== content) {
                yield* blockChildren(content);
            }

            continue;
        }

        yield child;
    }
}

/**
 * What a built-in HEADING is spaced by when nothing in the document says.
 *
 * Twelve points above it and six below, measured against LibreOffice with a
 * control beside it: two stub heading styles of the same size, one
 * silent and one stating `w:spacing` of zero on both sides, came out 24.50 and
 * 12.50 apart from the paragraph above, and 17.70 and 11.70 from the one below.
 * The difference is the whole of it — 12.00 and 6.00 — and it is what put every
 * heading in an under-specified document 12pt high and every paragraph after
 * one 6pt high.
 *
 * ## A LAST resort, which is what makes it safe
 *
 * It fires only where the merged properties state no spacing at all — not the
 * style, not `docDefaults`, not the paragraph itself. Measured, and this is the
 * half that matters: with a real `Normal` stating spacing of its own, a style
 * named `heading 8` came out at EXACTLY the same distance as a style named
 * `Custom Thing`, so the built-in does not fire when anything else answers.
 * An ordinary Word document, whose Normal always carries spacing, never reaches
 * this at all.
 *
 * Only the spacing. A real heading style is also bold, larger and kept with the
 * paragraph after it — none of which was measured here, because the stubs in
 * the probe stated their own size and weight, so none of it is invented here.
 */
function headingSpacing(
    styleId: string | null,
    styles: StyleSheet,
    properties: ParagraphProperties,
): ParagraphProperties {
    if (undefined !== properties.spaceBeforeTwips || undefined !== properties.spaceAfterTwips
        || !styles.isBuiltInHeading(styleId)) {
        return properties;
    }

    return {
        ...properties,
        spaceBeforeTwips: HEADING_SPACE_BEFORE_TWIPS,
        spaceAfterTwips: HEADING_SPACE_AFTER_TWIPS,
    };
}

/**
 * Twelve points and six, in the twips the rest of the reader speaks.
 *
 * Measured against LibreOffice as the DIFFERENCE between two stub
 * heading styles of the same size, one silent about `w:spacing` and one stating
 * zero on both sides: 24.50 against 12.50 above, and 17.70 against 11.70 below.
 */
const HEADING_SPACE_BEFORE_TWIPS = 240;
const HEADING_SPACE_AFTER_TWIPS = 120;

/**
 * `w:tblLook` — which of a table style's conditional formats are switched on.
 *
 * The HEX MASK is what answers, not the attributes beside it. Measured: a
 * table stating `w:firstRow="0"` alone kept its header dress, and the same
 * table stating `w:val="0000"` lost it. Word writes both and they can
 * disagree; the mask is the one that bites.
 *
 * The two banding bits are NEGATIVE — bands are on unless the mask says no.
 *
 * An ABSENT `w:tblLook` turns everything on, which is measured rather than
 * assumed: a table with no such element at all, over a style
 * dressing its first row and its odd bands, came out with BOTH — the header
 * fill on row one and the band fill on row two. Answering nought here, which
 * is the tempting reading of an absent mask, would have dressed neither.
 */
function tableLook(element: XmlElement | null): number {
    const stated = element?.attribute('w:val') ?? null;
    const parsed = null === stated ? Number.NaN : Number.parseInt(stated, 16);

    return Number.isNaN(parsed) ? LOOK_ALL : parsed;
}

/** Every positive condition, with the negative banding bits left clear. */
const LOOK_ALL = 0x01e0;

const LOOK_FIRST_ROW = 0x0020;
const LOOK_LAST_ROW = 0x0040;
const LOOK_FIRST_COLUMN = 0x0080;
const LOOK_LAST_COLUMN = 0x0100;
const LOOK_NO_H_BAND = 0x0200;
const LOOK_NO_V_BAND = 0x0400;

/** Where a cell sits, which is all the conditions are chosen by. */
interface CellPlace {
    readonly row: number;
    readonly rowCount: number;
    readonly column: number;
    readonly columnCount: number;
}

/**
 * The conditional formats dressing one cell, WEAKEST FIRST.
 *
 * Measured against LibreOffice, one table per condition so nothing
 * overlapped: `firstRow` filled row one across every column, `lastRow` the
 * last row, `firstCol` column one down every row, `lastCol` the last column,
 * and the two bands took alternate rows starting with `band1Horz`.
 *
 * ## Which wins where they meet
 *
 * A seventh table put a header, a first column and a band together and shaded
 * each differently. Row one came out the header's colour in EVERY column, and
 * row two came out the column's colour in the first cell and the band's in the
 * rest. So the order is bands, then the column conditions, then the row ones —
 * returned weakest first, for a caller that merges in order.
 *
 * The bands count the BODY, not the table: with a header on, row two was
 * `band1Horz` where the same table without one made row one band1.
 *
 * ## The bands that run the other way
 *
 * `band1Vert` and `band2Vert` do the same down the COLUMNS, gated by the mask's
 * `noVBand` bit, and they BEAT the horizontal ones: a table defining both, with
 * every cell in one of each, came out entirely the vertical band's colour.
 * That is the opposite of the order ECMA-376 lists, which is why it is
 * measured here rather than read off the specification.
 *
 * The first column is off the vertical banding exactly as the first row is off
 * the horizontal, and the count restarts after it: with `firstColumn` on, the
 * columns came out first-column, band1, band2, band1.
 *
 * ## What is NOT built: how deep a band is
 *
 * `w:tblStyleRowBandSize` and `w:tblStyleColBandSize` say how many rows or
 * columns a band covers. LibreOffice ignores both — a style stating two banded
 * every OTHER row and every other column, exactly as a style stating nothing
 * does — so following the file here would put the preview out of step with the
 * renderer that prints it, with no reference to check the result against.
 * Measured and left, which is what this engine does with a number nobody draws.
 */
function conditionsFor(
    conditions: ReadonlyMap<string, ConditionalFormat>,
    look: number,
    place: CellPlace,
): ConditionalFormat[] {
    const firstRow = 0 !== (look & LOOK_FIRST_ROW) && 0 === place.row;
    const lastRow = 0 !== (look & LOOK_LAST_ROW) && place.row === place.rowCount - 1;
    const firstColumn = 0 !== (look & LOOK_FIRST_COLUMN) && 0 === place.column;
    const lastColumn = 0 !== (look & LOOK_LAST_COLUMN)
        && place.column === place.columnCount - 1;
    const bodyFrom = 0 !== (look & LOOK_FIRST_ROW) ? 1 : 0;
    const insideFrom = 0 !== (look & LOOK_FIRST_COLUMN) ? 1 : 0;
    const out: ConditionalFormat[] = [];

    const add = (type: string): void => {
        const format = conditions.get(type);
        if (undefined !== format) {
            out.push(format);
        }
    };

    if (0 === (look & LOOK_NO_H_BAND) && !firstRow && !lastRow) {
        add(0 === (place.row - bodyFrom) % 2 ? 'band1Horz' : 'band2Horz');
    }
    if (0 === (look & LOOK_NO_V_BAND) && !firstColumn && !lastColumn) {
        add(0 === (place.column - insideFrom) % 2 ? 'band1Vert' : 'band2Vert');
    }
    if (firstColumn) {
        add('firstCol');
    }
    if (lastColumn) {
        add('lastCol');
    }
    if (lastRow) {
        add('lastRow');
    }
    if (firstRow) {
        add('firstRow');
    }

    return out;
}

/**
 * Block-level children of a given name, with content controls flattened.
 *
 * Every walk over a table's rows and a row's cells goes through this, because a
 * table read twice must read the same both times. It was not: `readTable`
 * flattened controls and `columnWidths` did not, so the same table wrapped in a
 * repeating section came out with its rows and its WIDTHS disagreeing — a
 * nominal grid recovered from the cells in one pass and left at 100 twips in
 * the other, six pixels to a column.
 */
function* blockChildrenNamed(parent: XmlElement, name: string): Generator<XmlElement> {
    for (const child of blockChildren(parent)) {
        if (name === child.name) {
            yield child;
        }
    }
}

/** The first descendant with this local name, whatever its prefix. */
function findByLocalName(root: XmlElement, name: string): XmlElement | null {
    for (const child of root.elements()) {
        if (localNameOf(child.name) === name) {
            return child;
        }

        const found = findByLocalName(child, name);
        if (null !== found) {
            return found;
        }
    }

    return null;
}

/** One declaration out of a VML shape's `style`. */
function styleValue(style: string, property: string): string | null {
    for (const declaration of style.split(';')) {
        const at = declaration.indexOf(':');
        if (at < 0) {
            continue;
        }
        if (declaration.slice(0, at).trim().toLowerCase() === property) {
            return declaration.slice(at + 1).trim();
        }
    }

    return null;
}

/**
 * A CSS length in the units VML writes, as pixels.
 *
 * Word states a shape's size in POINTS, but a document converted from .doc or
 * written by another producer may say inches or centimetres — and all three
 * printed identically through LibreOffice, so all three have to mean the same
 * thing here. A bare number is pixels, which is what CSS says it is.
 */
const CSS_UNIT_PX: Readonly<Record<string, number>> = {
    px: 1,
    pt: 96 / 72,
    pc: 96 / 6,
    in: 96,
    cm: 96 / 2.54,
    mm: 96 / 25.4,
};

function cssLengthToPx(length: string | null): number | null {
    if (null === length) {
        return null;
    }

    const match = /^\s*(-?[\d.]+)\s*([a-z]*)\s*$/i.exec(length);
    if (null === match) {
        return null;
    }

    const value = Number(match[1]);
    const unit = (match[2] ?? '').toLowerCase();
    const scale = '' === unit ? 1 : CSS_UNIT_PX[unit];

    return Number.isFinite(value) && undefined !== scale ? value * scale : null;
}

function readShading(element: XmlElement | null): string | null {
    const fill = element?.attribute('w:fill') ?? null;

    return null === fill || 'auto' === fill ? null : `#${fill}`;
}

/**
 * A `wp:positionH` or `wp:positionV`.
 *
 * Either a `posOffset` from a named origin or an `align` against it. An origin
 * this engine does not know becomes `column`, which is the writing area — the
 * safest place to put a picture whose anchor cannot be honoured, since it is
 * where the text is.
 */
function readFloatPosition(element: XmlElement | null): FloatPosition {
    const relative = element?.attribute('relativeFrom') ?? '';
    const align = byLocalName(element ?? EMPTY_ELEMENT, 'align')?.text ?? null;
    const offset = byLocalName(element ?? EMPTY_ELEMENT, 'posOffset')?.text ?? null;

    return {
        relativeTo: RELATIVE_TO[relative] ?? 'column',
        offsetPx: null === offset ? 0 : emuToPx(Number(offset) || 0),
        ...(null === align ? {} : { align }),
    };
}

const RELATIVE_TO: Record<string, RelativeTo> = {
    column: 'column',
    character: 'column',
    margin: 'margin',
    leftMargin: 'margin',
    rightMargin: 'margin',
    insideMargin: 'margin',
    outsideMargin: 'margin',
    page: 'page',
    paragraph: 'paragraph',
    line: 'paragraph',
    topMargin: 'margin',
    bottomMargin: 'margin',
};

/**
 * Which wrap a float asks for.
 *
 * `wrapTight` and `wrapThrough` follow the picture's OUTLINE rather than its
 * box; both are read as `square`, which keeps the text clear of the picture but
 * further from it than Word would. `wrapTopAndBottom` is read as square too,
 * which leaves text beside a float Word would have pushed below it — stated
 * rather than silently approximated.
 */
function readWrap(anchor: XmlElement): WrapMode {
    if (null !== byLocalName(anchor, 'wrapNone')) {
        return 'none';
    }

    return 'square';
}

/**
 * The space between a text box's frame and the words inside it.
 *
 * The defaults are the format's own — a tenth of an inch across, a twentieth
 * down — and they are exactly what LibreOffice printed for a box that stated
 * nothing: its text started 7.20 in from the frame and its first baseline 3.60
 * below the top, against 0 and 0 for the same box stating `lIns="0" tIns="0"`.
 * So the file is believed, and this is what silence means.
 */
const DEFAULT_INSET_EMU = { left: 91440, top: 45720, right: 91440, bottom: 45720 };

/** DrawingML's `wps:bodyPr`, which says the inset in EMU on four attributes. */
function readBodyInset(bodyPr: XmlElement | null): BoxInset {
    return {
        leftPx: emuToPx(attributeNumber(bodyPr, 'lIns') ?? DEFAULT_INSET_EMU.left),
        topPx: emuToPx(attributeNumber(bodyPr, 'tIns') ?? DEFAULT_INSET_EMU.top),
        rightPx: emuToPx(attributeNumber(bodyPr, 'rIns') ?? DEFAULT_INSET_EMU.right),
        bottomPx: emuToPx(attributeNumber(bodyPr, 'bIns') ?? DEFAULT_INSET_EMU.bottom),
    };
}

/**
 * VML's `inset`, which says all four sides in one attribute as CSS lengths —
 * left, top, right, bottom — any of which may be left empty for the default.
 *
 * LibreOffice does NOT read this: a box stating `inset="0,0,0,0"` printed its
 * text in exactly the same place as one stating nothing, 4.25 in from the frame
 * on both axes, which is a LibreOffice default rather than anything in the
 * file. The file is believed here instead, because a renderer that ignores
 * what a document says is not a specification of what the document means.
 */
function readVmlInset(inset: string | null): BoxInset {
    const parts = (inset ?? '').split(',');
    const side = (index: number, fallbackEmu: number): number =>
        cssLengthToPx((parts[index] ?? '').trim() || null) ?? emuToPx(fallbackEmu);

    return {
        leftPx: side(0, DEFAULT_INSET_EMU.left),
        topPx: side(1, DEFAULT_INSET_EMU.top),
        rightPx: side(2, DEFAULT_INSET_EMU.right),
        bottomPx: side(3, DEFAULT_INSET_EMU.bottom),
    };
}

/** A childless element, so the position reader needs no null dance. */
const EMPTY_ELEMENT = XmlDocument.parse('<empty/>').root;

/** The first child with this local name, whatever namespace prefix it carries. */
function byLocalName(parent: XmlElement, local: string): XmlElement | null {
    for (const child of parent.elements()) {
        if (localNameOf(child.name) === local) {
            return child;
        }
    }

    return null;
}

/** Follow a chain of local names down from an element. */
function descendByLocalName(from: XmlElement, path: readonly string[]): XmlElement | null {
    let current: XmlElement | null = from;
    for (const local of path) {
        if (null === current) {
            return null;
        }
        current = byLocalName(current, local);
    }

    return current;
}

function localNameOf(name: string): string {
    return name.slice(name.indexOf(':') + 1);
}

/** Whether two page setups describe the same paper. */
function sameGeometry(a: PageGeometry, b: PageGeometry): boolean {
    return a.widthPx === b.widthPx
        && a.heightPx === b.heightPx
        && a.marginTopPx === b.marginTopPx
        && a.marginBottomPx === b.marginBottomPx
        && a.marginLeftPx === b.marginLeftPx
        && a.marginRightPx === b.marginRightPx;
}

class BodyReader {
    readonly diagnostics: Diagnostic[] = [];
    private readonly reported = new Set<string>();

    /**
     * Counters live on the READER, so the same document read twice produces
     * the same numbers. On the shared definitions they would keep counting.
     */
    private readonly counters = new NumberingCounters();

    /**
     * Floats seen while reading the current paragraph's runs.
     *
     * On the reader because the drawing is found deep inside a run and the
     * float belongs to the paragraph — the two are several frames apart, and
     * threading a collector through every run would put a parameter on the run
     * reader that has nothing to do with runs.
     */
    private pendingFloats: Float[] = [];

    /**
     * What the table being read lends to the paragraphs inside it.
     *
     * On the reader for the same reason the pending floats are: the table
     * style is named on the table and spent several frames down, inside each
     * cell's paragraphs. Saved and put back around every table, so a table
     * nested in a cell of another does not carry its parent's dress out with
     * it.
     */
    private tableStyleLends: { paragraph: ParagraphProperties; run: RunProperties } = {
        paragraph: {},
        run: {},
    };

    /**
     * The style each paragraph came from, and whether it asked for contextual
     * spacing — remembered only until the blocks around it are known.
     *
     * A `WeakMap` because the key is the block itself: nothing downstream needs
     * this, and putting a style id on the layout model to carry it would be a
     * field the editor seam then has to classify for no other purpose.
     */
    private readonly paragraphStyleIds = new WeakMap<Block, {
        styleId: string | null;
        contextual: boolean;
        /** `w:snapToGrid` — false where the paragraph asked to come off it. */
        snapToGrid: boolean;
    }>();

    constructor(
        private readonly styles: StyleSheet,
        private readonly numbering: Numbering,
        private readonly fonts: FontCatalogue,
        private readonly defaultTabPx: number = DEFAULT_TAB_PX,
        private readonly mediaById: Readonly<Record<string, ImageContent>> = {},
        private readonly decimalSymbol: string | null = null,
    ) {}

    /**
     * What each note is NUMBERED, and the order they were first referenced in.
     *
     * A note's number is its position among the references in the BODY, not its
     * `w:id` — Word writes ids in whatever order the notes were created, and a
     * document edited more than once has them out of order. Held on the reader
     * because the notes are read after the body, and by then the order the body
     * asked for them in is the only thing that decides it.
     */
    private readonly numbered = new Map<number, number>();

    /** Set while a note's own blocks are being read, for its `w:footnoteRef`. */
    private footnoteBeingRead: number | null = null;

    /** The same, for endnotes — which are numbered in a series of their own. */
    private readonly numberedEndnotes = new Map<number, number>();

    private endnoteBeingRead: number | null = null;

    endnoteNumber(id: number): number {
        const known = this.numberedEndnotes.get(id);
        if (undefined !== known) {
            return known;
        }

        const next = this.numberedEndnotes.size + 1;
        this.numberedEndnotes.set(id, next);

        return next;
    }

    referencedEndnotes(): number[] {
        return [...this.numberedEndnotes.keys()];
    }

    readEndnote(id: number, element: XmlElement): Block[] {
        this.endnoteBeingRead = id;
        try {
            return this.read(element);
        } finally {
            this.endnoteBeingRead = null;
        }
    }

    footnoteNumber(id: number): number {
        const known = this.numbered.get(id);
        if (undefined !== known) {
            return known;
        }

        const next = this.numbered.size + 1;
        this.numbered.set(id, next);

        return next;
    }

    /** The ids the body referenced, in the order it referenced them. */
    referencedFootnotes(): number[] {
        return [...this.numbered.keys()];
    }

    readFootnote(id: number, element: XmlElement): Block[] {
        this.footnoteBeingRead = id;
        try {
            return this.read(element);
        } finally {
            this.footnoteBeingRead = null;
        }
    }

    read(body: XmlElement): Block[] {
        return this.readSections(body).flatMap((section) => section.blocks);
    }

    /**
     * Split the body at its section breaks.
     *
     * A `w:sectPr` inside a paragraph's properties ends the section that
     * paragraph BELONGS TO — the properties describe what came before them,
     * not what follows. The body's own trailing `w:sectPr` describes the last
     * section, which is why a document with no breaks still has one.
     */
    readSections(body: XmlElement): { blocks: Block[]; sectPr: XmlElement | null }[] {
        const sections: { blocks: Block[]; sectPr: XmlElement | null }[] = [];
        let blocks: Block[] = [];

        for (const child of blockChildren(body)) {
            if ('w:p' === child.name) {
                blocks.push(...this.readParagraph(child));

                const sectPr = child.element('w:pPr')?.element('w:sectPr') ?? null;
                if (null !== sectPr) {
                    sections.push({ blocks: this.dropContextualSpacing(blocks), sectPr });
                    blocks = [];
                }
            } else if ('w:tbl' === child.name) {
                blocks.push(this.readTable(child));
            }
        }

        sections.push({
            blocks: this.dropContextualSpacing(blocks),
            sectPr: body.element('w:sectPr'),
        });

        return sections;
    }

    /**
     * A `w:tbl`.
     *
     * ## The grid is the STRUCTURE; `w:tcW` is often the width
     *
     * `w:tblGrid` says how many columns there are and `w:gridSpan` says how
     * many of them a cell covers — but the grid's widths are frequently
     * nominal. The Word-authored fixture here declares `w:w="100"` for both
     * of its columns and puts the real 4680 twips on every `w:tcW`; trusting
     * the grid gives columns a third of a character wide, every cell wraps to
     * one letter per line, and the document gains a page.
     *
     * So the grid supplies the shape and the cells refine it. Word does the
     * same thing — it recomputes the grid from the cells when it opens a file.
     */
    private readTable(element: XmlElement): Table {
        const properties = element.element('w:tblPr') ?? null;
        // A table authored in Word names a STYLE and leaves its borders and
        // cell margins to it — "Table Grid" is the default, and its rules are
        // the only thing making the table look like a table. Nothing here read
        // `w:tblStyle` at all, so every such table came out bare.
        //
        // The style chain first and the table's own properties last, because
        // each level answers for what it names.
        const tableStyleId = properties?.element('w:tblStyle')?.attribute('w:val') ?? null;
        const inherited = this.styles.tableStyleProperties(tableStyleId);
        // Restored at the end of this method, so a nested table gives its
        // parent's dress back rather than keeping it.
        const outerLends = this.tableStyleLends;
        this.tableStyleLends = this.styles.tableStyleCascade(tableStyleId);
        const conditions = this.styles.tableStyleConditions(tableStyleId);
        const look = tableLook(properties?.element('w:tblLook') ?? null);
        const chain = [...inherited, ...(null === properties ? [] : [properties])];
        const marginElements = chain.map((level) => level.element('w:tblCellMar') ?? null);
        /** A cell margin, from the last level of the chain that names that side. */
        const marginOf = (side: string): number | null => marginElements.reduce(
            (found: number | null, element) => marginTwips(element, side) ?? found, null);

        const gridTwips = (element.element('w:tblGrid')?.elements('w:gridCol') ?? [])
            .map((column) => attributeNumber(column, 'w:w') ?? 0);
        const preferred = readPreferredWidth(properties?.element('w:tblW') ?? null);
        const columnWidthsPx = this
            .columnWidths(element, gridTwips, preferred.twips)
            .map(twipsToPx);
        const alignment = readTableAlignment(properties?.element('w:jc') ?? null);
        // `w:tblInd` is kept whether or not it resolves to a distance:
        // LibreOffice still pulled a table left by its cell margin for a
        // percentage indent it could make nothing else of, so PRESENT and
        // unresolvable is nought, not absent.
        const indent = properties?.element('w:tblInd') ?? null;
        const indentTwips = 'dxa' === (indent?.attribute('w:type') ?? 'dxa')
            ? attributeNumber(indent, 'w:w') ?? 0
            : 0;

        const rows: TableRow[] = [];
        const rowElements = [...blockChildrenNamed(element, 'w:tr')];
        for (const row of rowElements) {
            const rowProperties = row.element('w:trPr');

            const cells: TableCell[] = [];
            for (const cell of blockChildrenNamed(row, 'w:tc')) {
                const cellProperties = cell.element('w:tcPr');
                // What the style lends THIS cell: the table's own dress, then
                // whichever conditional formats reach it. Set before the
                // paragraphs are read, because that is where it is spent, and
                // put back after the row so the next one starts clean.
                const applying = conditionsFor(conditions, look, {
                    row: rows.length,
                    rowCount: rowElements.length,
                    column: cells.length,
                    columnCount: gridTwips.length,
                });
                const lends = this.tableStyleLends;
                for (const format of applying) {
                    this.tableStyleLends = {
                        paragraph: { ...this.tableStyleLends.paragraph, ...format.paragraph },
                        run: { ...this.tableStyleLends.run, ...format.run },
                    };
                }
                // In document order, because a cell holds paragraphs and
                // tables interleaved — reading the paragraphs alone lost the
                // tables, and reading them separately would put them all at
                // one end.
                const paragraphs: Block[] = [];
                for (const child of blockChildren(cell)) {
                    if ('w:p' === child.name) {
                        paragraphs.push(...this.readParagraph(child));
                    } else if ('w:tbl' === child.name) {
                        paragraphs.push(this.readTable(child));
                    }
                }
                // A cell stacks paragraphs the way the body does, so a list
                // inside one is spaced by the same rule.
                paragraphs.splice(0, paragraphs.length, ...this.dropContextualSpacing(paragraphs));

                // `w:textDirection`: read here so the model is whole, and
                // laid out by nobody yet — see the field's own note.
                const turned = cellProperties?.element('w:textDirection')
                    ?.attribute('w:val') ?? null;
                const textDirection: CellTextDirection | null =
                    'btLr' === turned || 'tbRl' === turned ? turned : null;
                const cellMargins = readCellMargins(
                    cellProperties?.element('w:tcMar') ?? null,
                );
                this.tableStyleLends = lends;
                // A conditional format carries borders as well as shading —
                // a heavy rule under the header is the commonest table look
                // there is. Under the cell's OWN, per side, exactly as the
                // table style sits under the table: measured, a
                // style lending its first row a 3pt bottom drew 3.000 there
                // and 0.500 elsewhere, and the same style over a cell saying
                // `none` drew the table's 0.500 throughout.
                const cellBorders = readBordersFrom([
                    ...applying.map((format) => format.tcPr?.element('w:tcBorders') ?? null),
                    cellProperties?.element('w:tcBorders') ?? null,
                ]);
                // The cell's own shading first; a conditional format answers
                // only where the cell is silent.
                const cellFill = readShading(cellProperties?.element('w:shd') ?? null)
                    ?? applying.reduce(
                        (found: string | null, format) =>
                            readShading(format.tcPr?.element('w:shd') ?? null) ?? found,
                        null);

                const verticalAlign = readCellVerticalAlign(
                    cellProperties?.element('w:vAlign') ?? null,
                );
                const merge = cellProperties?.element('w:vMerge') ?? null;
                // `w:vMerge` with no value at all means `continue`, which is
                // how most files write it: the attribute is only ever there to
                // say `restart`.
                const verticalMerge: VerticalMerge | null = null === merge
                    ? null
                    : ('restart' === merge.attribute('w:val') ? 'restart' : 'continue');

                cells.push({
                    paragraphs,
                    gridSpan: attributeNumber(cellProperties?.element('w:gridSpan') ?? null, 'w:val') ?? 1,
                    ...(undefined === cellBorders ? {} : { borders: cellBorders }),
                    ...(undefined === cellMargins ? {} : { margins: cellMargins }),
                    ...(null === textDirection ? {} : { textDirection }),
                    ...(null === cellFill ? {} : { shadingFill: cellFill }),
                    ...(null === verticalAlign ? {} : { verticalAlign }),
                    ...(null === verticalMerge ? {} : { verticalMerge }),
                });
            }

            const height = rowProperties?.element('w:trHeight') ?? null;
            const heightTwips = attributeNumber(height, 'w:val');

            rows.push({
                cells,
                isHeader: true === toggleOf(rowProperties, 'w:tblHeader'),
                cantSplit: true === toggleOf(rowProperties, 'w:cantSplit'),
                ...(null === heightTwips ? {} : {
                    heightPx: twipsToPx(heightTwips),
                    heightRule: 'exact' === height?.attribute('w:hRule') ? 'exact' : 'atLeast',
                }),
            });
        }

        // Word's default cell padding is 0.08 inch either side and none top or
        // bottom. A table that declares nothing still has it, so text would
        // otherwise sit against the borders and every column would be one
        // character wider than the file says.
        this.tableStyleLends = outerLends;
        const tableBorders = readBordersFrom(
            chain.map((level) => level.element('w:tblBorders') ?? null));
        const tableFill = readShading(properties?.element('w:shd') ?? null);

        return {
            rows,
            columnWidthsPx,
            ...(null === preferred.fraction ? {} : { preferredWidthFraction: preferred.fraction }),
            ...(null === preferred.twips ? {} : { preferredWidthPx: twipsToPx(preferred.twips) }),
            ...(null === alignment ? {} : { alignment }),
            ...(null === indent ? {} : { indentPx: twipsToPx(indentTwips) }),
            ...(undefined === tableBorders ? {} : { borders: tableBorders }),
            ...(null === tableFill ? {} : { shadingFill: tableFill }),
            cellMarginLeftPx: twipsToPx(marginOf('w:left') ?? 115),
            cellMarginRightPx: twipsToPx(marginOf('w:right') ?? 115),
            cellMarginTopPx: twipsToPx(marginOf('w:top') ?? 0),
            cellMarginBottomPx: twipsToPx(marginOf('w:bottom') ?? 0),
            spaceBeforePx: 0,
            spaceAfterPx: 0,
            pageBreakBefore: false,
        };
    }

    /**
     * Effective column widths in twips.
     *
     * Spanning cells are apportioned FIRST and single-column cells applied
     * after, so an exact width always beats a share of a span — a row of
     * ordinary cells knows its columns better than a merged one above it.
     */
    private columnWidths(
        table: XmlElement,
        gridTwips: readonly number[],
        tableTwips: number | null,
    ): number[] {
        const widths = [...gridTwips];
        const spanning: { column: number; span: number; value: number }[] = [];
        const single: { column: number; value: number }[] = [];

        for (const row of blockChildrenNamed(table, 'w:tr')) {
            let column = 0;
            for (const cell of blockChildrenNamed(row, 'w:tc')) {
                const cellProperties = cell.element('w:tcPr');
                const span = Math.max(
                    1,
                    attributeNumber(cellProperties?.element('w:gridSpan') ?? null, 'w:val') ?? 1,
                );
                const declared = cellProperties?.element('w:tcW') ?? null;
                const value = attributeNumber(declared, 'w:w');
                const type = declared?.attribute('w:type') ?? 'dxa';

                // A percentage of the TABLE, not of the page: LibreOffice
                // gave two fifty-percent columns 150pt each inside a table
                // declared 6000 twips wide, which is half of 300pt.
                const resolved = 'pct' === type && null !== value && null !== tableTwips
                    ? Math.round(tableTwips * value / PERCENT_UNITS)
                    : value;

                if (null !== declared && null !== value && 'dxa' !== type
                    && null === tableTwips) {
                    // Nothing to resolve the percentage against. LibreOffice
                    // does the same — fifty-percent columns in a table of no
                    // declared width printed at their GRID widths, untouched.
                    this.report(
                        'unsupported-block',
                        `a table column width of type "${type}" fell back to the grid`,
                    );
                } else if (null !== resolved && resolved > 0) {
                    if (1 === span) {
                        single.push({ column, value: resolved });
                    } else {
                        spanning.push({ column, span, value: resolved });
                    }
                }

                column += span;
            }
        }

        for (const { column, span, value } of spanning) {
            const covered = widths.slice(column, column + span);
            const total = covered.reduce((sum, width) => sum + width, 0);
            for (let index = 0; index < span; index++) {
                // Proportionally when the grid has usable ratios, evenly when
                // it is the nominal all-equal kind.
                widths[column + index] = total > 0
                    ? (value * (covered[index] ?? 0)) / total
                    : value / span;
            }
        }

        for (const { column, value } of single) {
            widths[column] = value;
        }

        return widths;
    }

    /**
     * One `w:p` becomes one paragraph — or SEVERAL, when it contains explicit
     * page breaks.
     *
     * `<w:br w:type="page"/>` sits inside a run, part-way through a paragraph.
     * The layout model puts page breaks on paragraph boundaries, so the
     * paragraph is split there and the remainder starts a new page. That is
     * exactly what the break means, and it needs nothing new from the layout.
     *
     * ## An EMPTY piece of that split draws nothing
     *
     * A whole paragraph with no runs is a line — its mark has formatting and
     * occupies one. A PIECE of a split has no mark of its own, and measured
     * against LibreOffice it takes no room at all. Four arrangements,
     * every following page opening at the same 760.49:
     *
     *   * `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` — both pieces empty,
     *     and the paragraph draws nothing on either page
     *   * text then a break — the empty second piece draws nothing
     *   * a break then text — the empty first piece draws nothing
     *   * text, break, text — both drawn, which is the control
     *
     * That first one is how PHPWord writes `addPageBreak()`, so it is in every
     * document this platform generates: each opened with a blank line at the
     * top of every page after the first, and the whole page sat 11.50 low.
     *
     * Left with NO runs rather than dropped, which is what keeps the break: the
     * flow gives a run-less paragraph no line and still honours the
     * `pageBreakBefore` every piece after the first carries. Dropping the piece
     * would take the break with it and run the pages together.
     */
    private readParagraph(element: XmlElement): Paragraph[] {
        const pPr = element.element('w:pPr');
        const styleId = pPr?.element('w:pStyle')?.attribute('w:val') ?? null;

        if (null !== styleId && !this.styles.hasParagraphStyle(styleId)) {
            this.report('unknown-paragraph-style', `w:pStyle "${styleId}" is not defined in styles.xml`);
        }
        // docDefaults, then the style, then NUMBERING, then direct
        // formatting — the order WordprocessingML specifies. A list's indent
        // comes from its numbering level and a paragraph may still override
        // it, which is how one bullet gets pushed further in than its
        // neighbours.
        const list = this.readNumbering(pPr, styleId);
        const merged: ParagraphProperties = headingSpacing(styleId, this.styles, {
            ...this.styles.paragraphProperties(styleId, this.tableStyleLends.paragraph),
            ...list.properties,
            ...readDirectParagraphProperties(pPr),
        });

        // The marker is placed once the indents have been merged, and it takes
        // the first line's indent with it: for a list paragraph the hanging
        // indent belongs to the MARKER, and the text starts where the suffix
        // tab lands.
        const placed = null === list.marker ? null : this.placeMarker(list.marker, merged);
        const properties: ParagraphProperties = null === placed
            ? merged
            : { ...merged, indentFirstLineTwips: placed.firstLineTwips };

        // Floats are gathered while the runs are read, because that is where
        // the drawing is — but they belong to the PARAGRAPH, so they are taken
        // off the reader once its runs are done.
        this.pendingFloats = [];
        const segments = this.readRuns(element, styleId);
        const floats = this.pendingFloats;
        this.pendingFloats = [];
        const fallbackFont = this.resolveRun(this.styles.runProperties(styleId, null, this.tableStyleLends.run));

        const contextual = true === properties.contextualSpacing;
        const snapToGrid = false !== properties.snapToGrid;

        return segments.map((runs, index) => {
            // An empty paragraph still occupies a line, and its height comes
            // from the formatting on its paragraph mark. With no run at all the
            // line would be zero high and the page would hold extra content.
            //
            // A PIECE of a paragraph split by a page break is the exception and
            // keeps nothing: see the note above this method.
            const filled = 0 === runs.length && 1 === segments.length
                ? [{ ...fallbackFont, text: '' }]
                : runs;

            return {
                runs: filled,
                style: this.paragraphStyle(properties, filled, {
                    first: 0 === index,
                    last: index === segments.length - 1,
                }),
                // Only the first piece of a paragraph split by a page break
                // carries the marker: the continuation is the same list item,
                // and numbering it again would invent an entry.
                ...(0 === index && null !== placed ? { marker: placed.marker } : {}),
                // The floats go with the first piece for the same reason: they
                // are anchored to the paragraph, and a continuation on the next
                // page would place them a second time.
                ...(0 === index && 0 !== floats.length ? { floats } : {}),
            };
        }).map((paragraph) => {
            // What `w:contextualSpacing` needs is the STYLE of the neighbour,
            // which the layout model has no reason to carry and this reader
            // will not know again once the blocks are handed on. So it is
            // remembered here, against the block itself, and spent by
            // `dropContextualSpacing` as soon as the neighbours are known.
            this.paragraphStyleIds.set(paragraph, { styleId, contextual, snapToGrid });

            return paragraph;
        });
    }

    /**
     * Zero the space a `w:contextualSpacing` paragraph would spend against a
     * neighbour of its OWN style.
     *
     * Measured against LibreOffice: a pair 10pt apart either side
     * printed 21.50 without the flag and 11.50 with it, and the space went
     * against a plain neighbour of the same style as readily as against
     * another flagged one — so the flag belongs to the paragraph giving up its
     * own space, and it is asked of each side separately.
     *
     * Done here rather than in the layout because this is the last place that
     * knows which style each paragraph came from. A list is the case it exists
     * for: Word's list styles set it, so every list in every document is spaced
     * like a run of ordinary paragraphs until this runs.
     */
    /**
     * Whether this block asked to stay ON the section's grid.
     *
     * True for anything the reader has no note about — a table, or a block from
     * somewhere else — because the grid is the default and `w:snapToGrid` is
     * the exception.
     */
    snapsToGrid(block: Block): boolean {
        return false !== this.paragraphStyleIds.get(block)?.snapToGrid;
    }

    private dropContextualSpacing(blocks: readonly Block[]): Block[] {
        const sameStyle = (a: Block | undefined, b: Block | undefined): boolean => {
            if (undefined === a || undefined === b || isTable(a) || isTable(b)) {
                return false;
            }
            const left = this.paragraphStyleIds.get(a);
            const right = this.paragraphStyleIds.get(b);

            return undefined !== left && undefined !== right && left.styleId === right.styleId;
        };

        return blocks.map((block, index) => {
            const known = isTable(block) ? undefined : this.paragraphStyleIds.get(block);
            if (isTable(block) || undefined === known || !known.contextual) {
                return block;
            }

            const before = sameStyle(blocks[index - 1], block) ? 0 : block.style.spaceBeforePx;
            const after = sameStyle(block, blocks[index + 1]) ? 0 : block.style.spaceAfterPx;

            return before === block.style.spaceBeforePx && after === block.style.spaceAfterPx
                ? block
                : { ...block, style: { ...block.style, spaceBeforePx: before, spaceAfterPx: after } };
        });
    }

    /**
     * A paragraph's list membership: the indents its level dictates, and the
     * marker to draw beside its first line.
     *
     * The counter is advanced HERE, once per paragraph in reading order, which
     * is the only place the document's order is known.
     */
    private readNumbering(
        pPr: XmlElement | null,
        styleId: string | null,
    ): { properties: ParagraphProperties; marker: PendingMarker | null } {
        // A paragraph may name a list itself, or take one from its STYLE —
        // which is how Word's own "List Paragraph" works, and a reader that
        // looked only at the paragraph would see no list at all.
        const numPr = pPr?.element('w:numPr') ?? null;
        const own = numPr?.element('w:numId')?.attribute('w:val') ?? null;
        const fromStyle = this.styles.numbering(styleId);

        const numId = own ?? fromStyle?.numId ?? null;
        if (null === numId || '0' === numId) {
            // numId 0 is how a paragraph says it is NOT in a list, overriding a
            // style that puts it in one.
            return { properties: {}, marker: null };
        }

        const declared = numPr?.element('w:ilvl')?.attribute('w:val') ?? null;
        const ilvl = null !== declared
            ? Number(declared)
            : (null === own ? fromStyle?.ilvl ?? 0 : 0);
        const level = this.numbering.level(numId, Number.isFinite(ilvl) ? ilvl : 0);
        if (null === level) {
            this.report('unknown-numbering', `w:numId "${numId}" is not defined in numbering.xml`);

            return { properties: {}, marker: null };
        }

        // `w:lvlRestart` above zero names the level whose change restarts this
        // one; this engine restarts under ANY shallower level. It
        // cannot be measured — LibreOffice ignores the attribute in every
        // spelling — so it is reported rather than left silent.
        if (null !== level.restart && level.restart > 0) {
            this.report(
                'unknown-numbering',
                `w:lvlRestart "${level.restart}" is read as restarting under any level above`,
            );
        }

        const text = this.counters.next(this.numbering, numId, Number.isFinite(ilvl) ? ilvl : 0);
        const run = this.styles.runProperties(styleId, null, this.tableStyleLends.run);
        const resolved = this.fonts.resolve(
            level.fontFamily ?? run.family ?? null,
            true === run.bold,
            true === run.italic,
        );
        const sizePx = undefined === run.sizeHalfPoints
            ? halfPointsToPx(this.fonts.defaultSizePt * 2)
            : halfPointsToPx(run.sizeHalfPoints);

        if ('' === text) {
            return { properties: levelIndents(level), marker: null };
        }

        return {
            // The level's `w:ind` is offered exactly as a paragraph's own
            // would be, hanging and all, so the ordinary merge decides which
            // wins. It has to: a paragraph that states `w:ind` of its own is
            // laid out entirely by ITS numbers, marker included.
            properties: levelIndents(level),
            marker: {
                run: { text, font: resolved.font, sizePx, ...(true === run.kerned ? { kerned: true } : {}) },
                // Through `advanceOf` so the marker is measured by the same
                // rule as the text beside it — kerned only where the run says
                // `w:kern`. A right-justified marker is placed BY
                // this width, so the two must not disagree.
                widthPx: advanceOf(text, { font: resolved.font, sizePx, kerned: true === run.kerned }),
                spacePx: 'space' === level.suffix
                    ? advanceOf(' ', { font: resolved.font, sizePx, kerned: true === run.kerned })
                    : 0,
                justification: level.justification,
                suffix: level.suffix,
            },
        };
    }

    /**
     * Where a list marker is drawn, and where the text after it starts.
     *
     * Deferred until the paragraph's properties are merged because both
     * answers need the EFFECTIVE indents, and a paragraph may state its own —
     * in which case the level's are not consulted at all (a level
     * saying 720/360 under a paragraph saying 1000/200 printed its marker at
     * the paragraph's numbers).
     *
     * Two rules, both measured off the printed page:
     *
     * - The marker is justified around the ANCHOR, `indentLeft - hanging`.
     *   `left` starts there, `right` ends there — growing away from the text,
     *   never into it — and `center` straddles it. The hanging indent is
     *   therefore not a first-line indent at all: it positions the marker, and
     *   the text of the first line is decided separately, below.
     * - A `tab` suffix — Word's default — takes the text to the first tab stop
     *   at or after the marker's RIGHT edge. The stops are the paragraph's
     *   explicit `w:tabs` where it has any, its own left indent, and the
     *   repeating default; all measured from the margin.
     *
     * A `space` or `nothing` suffix has no stop to find, so the text follows
     * the marker directly and is never pulled left of the indent. That much is
     * unmeasured — no printed page here separates it from the tab case — and
     * is what this did for every suffix before the tab rule was measured.
     */
    private placeMarker(
        marker: PendingMarker,
        properties: ParagraphProperties,
    ): { marker: ListMarker; firstLineTwips: number } {
        const indentLeftPx = twipsToPx(properties.indentLeftTwips ?? 0);
        // A hanging indent is a NEGATIVE first-line indent; a positive one is
        // an ordinary first-line indent and gives the marker no room at all.
        const hangingPx = Math.max(0, -twipsToPx(properties.indentFirstLineTwips ?? 0));
        const anchorPx = indentLeftPx - hangingPx;
        const shiftPx = 'right' === marker.justification
            ? marker.widthPx
            : ('center' === marker.justification ? marker.widthPx / 2 : 0);

        const leftPx = anchorPx - shiftPx;
        const rightPx = leftPx + marker.widthPx;
        const textPx = 'tab' === marker.suffix
            ? this.firstTabStopAt(rightPx, properties, indentLeftPx)
            : Math.max(indentLeftPx, rightPx + marker.spacePx);

        return {
            marker: { run: marker.run, offsetPx: indentLeftPx - leftPx },
            firstLineTwips: pointsToTwips(pxToPoints(textPx - indentLeftPx)),
        };
    }

    /**
     * The first tab stop at or after `xPx`, measured from the left margin.
     *
     * The paragraph's own left indent is a stop (a marker ending
     * before an indent of 25pt put its text at 25pt, not at the 36pt default
     * stop past it). Explicit stops suppress the defaults below them, which is
     * the rule the line breaker already follows for a tab in the text.
     */
    private firstTabStopAt(
        xPx: number,
        properties: ParagraphProperties,
        indentLeftPx: number,
    ): number {
        const explicit = (properties.tabStops ?? [])
            .map((stop) => twipsToPx(stop.positionTwips))
            .filter((stop) => stop >= xPx);
        // AT or after, not strictly after: a marker whose right edge lands
        // exactly on a stop is a case no page here prints, and advancing past
        // it would be a guess in the other direction.
        const beyond = explicit.length > 0
            ? Math.min(...explicit)
            : Math.ceil(xPx / this.defaultTabPx) * this.defaultTabPx;

        return indentLeftPx >= xPx ? Math.min(indentLeftPx, beyond) : beyond;
    }

    /**
     * The runs of a paragraph, grouped into the segments a page break divides
     * it into.
     */
    private readRuns(paragraph: XmlElement, styleId: string | null): StyledRun[][] {
        const segments: StyledRun[][] = [[]];

        // The complex form is a RUN SEQUENCE, not a container: a `begin`
        // fldChar, the instruction in `w:instrText`, a `separate`, the value
        // Word last computed, then `end`. Only what lies between `separate` and
        // `end` is the result, so the state has to survive across siblings —
        // which is why it is threaded into the run reader rather than kept
        // inside it.
        const state: FieldState = {
            kind: null,
            format: null,
            instruction: null,
            inResult: false,
        };

        const visit = (parent: XmlElement): void => {
            for (const child of parent.elements()) {
                switch (child.name) {
                    case 'w:pPr':
                        break;
                    case 'w:del':
                    case 'w:moveFrom':
                        // `w:moveFrom` is the half of a tracked MOVE the text
                        // left behind, and it disappeared here already — but by
                        // accident, because its runs hold `w:delText` and only
                        // `w:t` is read. Named outright so the two do not drift
                        // apart, and so the notice below covers both.
                        //
                        // Text deleted under revision tracking. It is still in
                        // the file, and this engine draws the document as it
                        // would READ once the changes are accepted — the text
                        // without its markup.
                        //
                        // MEASURED, NOT BUILT. LibreOffice prints the
                        // other view, the one a word processor shows an editor:
                        // `DELETED` drawn in place, struck through by a 0.60pt
                        // rule 44.90 long sitting 2.60 above the baseline, with
                        // a change bar down the margin at 64.95 — 7.15pt left
                        // of the text. So it is markup, not content, and a page
                        // that shows it is answering a different question.
                        //
                        // Reported, because dropping text silently is the one
                        // thing a reader must not do — the same rule that
                        // covers a character it cannot draw.
                        this.report(
                            'revision-hidden',
                            'text deleted under revision tracking is not drawn',
                        );
                        break;
                    case 'w:r':
                        this.readRun(child, styleId, segments, state);
                        break;
                    case 'w:fldSimple': {
                        // The simple form carries its instruction as an
                        // attribute and its cached value as ordinary runs
                        // inside, so descending with the kind set is what marks
                        // those runs as the field's result.
                        const outer = { ...state };
                        state.kind = fieldKind(child.attribute('w:instr'));
                        state.format = fieldFormat(child.attribute('w:instr'));
                        state.inResult = true;
                        visit(child);
                        Object.assign(state, outer);
                        break;
                    }
                    default:
                        // w:hyperlink, w:ins, w:smartTag, w:bookmarkStart and
                        // friends wrap runs without changing them. Descending
                        // is what keeps a hyperlink's text in the document.
                        visit(child);
                }
            }
        };
        visit(paragraph);

        return segments;
    }

    private readRun(
        run: XmlElement,
        styleId: string | null,
        segments: StyledRun[][],
        state: FieldState,
    ): void {
        const characterStyle = run.element('w:rPr')?.element('w:rStyle')?.attribute('w:val') ?? null;
        const properties: RunProperties = {
            ...this.styles.runProperties(styleId, characterStyle, this.tableStyleLends.run),
            ...readDirectRunProperties(run.element('w:rPr')),
        };

        // `w:vanish` — the document says to hide this run, so it does not reach
        // the layout at all. Measured: LibreOffice prints neither the
        // text nor the room it would take, the words either side of a hidden
        // one sitting exactly as far apart as the first one is wide. Returning
        // here rather than dropping the ink later is what makes the width go
        // too — and what keeps hidden text out of a rendered page, which is
        // the whole point of the property.
        if (true === properties.hidden) {
            return;
        }

        const resolved = this.resolveRun(properties);

        // `w:caps` draws the run in capitals whatever the file stores. Applied
        // to the TEXT rather than at drawing time because it changes the width:
        // measured in lower case, a capitalised run breaks its line early and
        // the page count follows.
        const capitalise = true === properties.caps;
        // `w:smallCaps` capitalises too, but only the LOWER case is made
        // smaller — measured against LibreOffice, where a capital in a small
        // caps run kept its full size and the letters beside it did not.
        const smallCaps = !capitalise && true === properties.smallCaps;

        let text = '';
        const flush = (): void => {
            if ('' !== text) {
                const kind = state.inResult ? state.kind : null;
                const numerals = state.inResult ? state.format : null;
                const extra = {
                    ...(null === kind ? {} : { field: kind }),
                    ...(null === numerals ? {} : { fieldFormat: numerals }),
                };

                for (const piece of smallCaps ? smallCapped(text) : [{ text, small: false }]) {
                    segments[segments.length - 1]!.push({
                        ...resolved,
                        text: capitalise || piece.small ? piece.text.toUpperCase() : piece.text,
                        ...(piece.small
                            ? { sizePx: resolved.sizePx * SMALL_CAPS_SCALE }
                            : {}),
                        ...extra,
                    });
                }
                text = '';
            }
        };

        for (const child of run.elements()) {
            switch (child.name) {
                case 'w:t':
                    text += child.text;
                    break;
                case 'w:tab':
                    // Carried as the character, and the line breaker turns it
                    // into a distance: a tab's width is the way to its next
                    // STOP, which is not known until the line has got there.
                    // (This said "tab stops are not implemented" long after
                    // they were — `w:tabs` reaches `style.tabStops`, defaults
                    // and decimal alignment included.)
                    text += '\t';
                    break;
                case 'w:noBreakHyphen':
                    text += '‑';
                    break;
                case 'w:sym': {
                    // A character named by CODE, in a font of its own. Word
                    // writes bullets and dingbats this way, and a reader that
                    // passed over it dropped them without a word.
                    const point = symbolCodePoint(child.attribute('w:char'));
                    const family = child.attribute('w:font');
                    if (null === point) {
                        break;
                    }

                    const face = this.fonts.resolve(family, false, false);
                    if (null !== face.substitutedFor) {
                        // The glyph is only in the font the document names, and
                        // that font is not here. Drawing the code point out of
                        // a substitute prints whatever happens to live there —
                        // an `a` where the document wanted an arrow.
                        this.report(
                            'font-substituted',
                            `a w:sym in "${face.substitutedFor}" was dropped; `
                                + 'the font is not available',
                        );
                        break;
                    }

                    flush();
                    segments[segments.length - 1]!.push({
                        ...resolved,
                        font: face.font,
                        text: String.fromCodePoint(point),
                    });
                    break;
                }
                case 'w:endnoteReference':
                case 'w:endnoteRef': {
                    // The same shape as a footnote's mark, and numbered the
                    // way Word numbers endnotes: measured against LibreOffice,
                    // the first two came out `i` and `ii`.
                    flush();
                    const id = 'w:endnoteRef' === child.name
                        ? this.endnoteBeingRead
                        : attributeNumber(child, 'w:id');
                    if (null === id) {
                        break;
                    }

                    // Word's built-in `EndnoteReference` raises the mark, and
                    // LibreOffice draws it raised whether or not the file
                    // carries the style — so the default is applied here for
                    // the same reason the numbering is. Only where the
                    // document has not raised it ALREADY, or the two would be
                    // applied one on top of the other.
                    // The mark in the TEXT only: the copy that opens the note
                    // sits on its own baseline, which is where LibreOffice
                    // drew it.
                    const raised = 'w:endnoteReference' === child.name
                        && undefined === resolved.baselineShiftPx
                        ? {
                            sizePx: resolved.sizePx * SCRIPT_SCALE,
                            baselineShiftPx: -resolved.sizePx * SUPERSCRIPT_RISE,
                        }
                        : {};

                    segments[segments.length - 1]!.push({
                        ...resolved,
                        ...raised,
                        text: formatNumeral(this.endnoteNumber(id), ENDNOTE_NUMERALS),
                    });
                    break;
                }
                case 'w:footnoteReference':
                case 'w:footnoteRef': {
                    // The mark ends whatever text preceded it in this run, and
                    // then IS a run of its own — its number is not part of the
                    // sentence and must be able to carry the superscript the
                    // `FootnoteReference` style puts on it.
                    flush();
                    const id = 'w:footnoteRef' === child.name
                        ? this.footnoteBeingRead
                        : attributeNumber(child, 'w:id');
                    if (null === id) {
                        break;
                    }

                    const number = this.footnoteNumber(id);
                    segments[segments.length - 1]!.push({
                        ...resolved,
                        text: String(number),
                        // Only a reference in the BODY reserves room; the copy
                        // that opens the note itself is already down there.
                        ...('w:footnoteReference' === child.name ? { footnoteId: id } : {}),
                    });
                    break;
                }
                case 'w:fldChar':
                    // Flush FIRST: text before a boundary belongs to whichever
                    // side of it the run has been on so far.
                    flush();
                    switch (child.attribute('w:fldCharType')) {
                        case 'begin':
                            state.instruction = '';
                            state.inResult = false;
                            break;
                        case 'separate':
                            state.kind = fieldKind(state.instruction);
                            state.format = fieldFormat(state.instruction);
                            state.inResult = true;
                            break;
                        case 'end':
                            state.kind = null;
                            state.format = null;
                            state.instruction = null;
                            state.inResult = false;
                            break;
                        default:
                            break;
                    }
                    break;
                case 'w:instrText':
                    if (null !== state.instruction && !state.inResult) {
                        state.instruction += child.text;
                    }
                    break;
                case 'w:drawing': {
                    // A picture ends whatever text preceded it in this run: as
                    // far as the line is concerned they are two different runs.
                    flush();
                    // Before the picture readers, both of them: a text box has
                    // no `a:blip` in it, so `readAnchor` would report a
                    // floating picture with no image and this would come out
                    // as a fault in a file that has none.
                    const box = this.readTextBox(child, 'DrawingML');
                    if (null !== box) {
                        this.pendingFloats.push(box);
                        break;
                    }
                    // A box that sits IN the line rather than beside it: not a
                    // float, and not a fault either.
                    const inlineBox = this.readInlineTextBox(child);
                    if (null !== inlineBox) {
                        segments[segments.length - 1]!.push({
                            ...resolved,
                            text: OBJECT_REPLACEMENT,
                            shape: inlineBox,
                        });
                        break;
                    }
                    const anchored = this.readAnchor(child);
                    if (null !== anchored) {
                        // A float belongs to the PARAGRAPH, not to this run —
                        // it is anchored to the paragraph and placed against
                        // it, and it occupies no width in the line at all.
                        this.pendingFloats.push(anchored);
                        break;
                    }
                    const image = this.readDrawing(child);
                    if (null !== image) {
                        segments[segments.length - 1]!.push({
                            ...resolved,
                            text: OBJECT_REPLACEMENT,
                            image,
                        });
                    }
                    break;
                }
                case 'w:object': {
                    // An embedded object — a chart, an equation, a spreadsheet
                    // — which Word stores beside a VML PICTURE of what it looks
                    // like. The object itself is not something this can run;
                    // the picture is what a reader sees, and dropping it lost
                    // the whole thing.
                    flush();
                    const drawn = this.readPicture(child);
                    if (null !== drawn) {
                        segments[segments.length - 1]!.push({
                            ...resolved,
                            text: OBJECT_REPLACEMENT,
                            image: drawn,
                        });
                    }
                    break;
                }
                case 'w:ruby': {
                    // Furigana and its like: a BASE of ordinary text with a
                    // gloss set above it. The base is the sentence itself, so
                    // passing the whole element over took the sentence with it.
                    const base = byLocalName(child, 'rubyBase');
                    if (null === base) {
                        break;
                    }

                    // The base is read as runs of its own, which reach the
                    // paragraph at once — so whatever text this run has
                    // buffered has to reach it first or the two come out in
                    // the wrong order.
                    flush();

                    const before = segments[segments.length - 1]!.length;
                    for (const inner of base.elements('w:r')) {
                        this.readRun(inner, styleId, segments, state);
                    }

                    // The gloss rides on the FIRST run of the base, which is
                    // the one it is set over. A base of several runs — a word
                    // that changes style half way — would want it spread
                    // across them, and no measurement covers that; the report
                    // says so rather than the gloss going quietly.
                    const gloss = this.readRubyGloss(child, styleId);
                    const carrier = segments[segments.length - 1]![before];
                    if (undefined !== gloss && undefined !== carrier) {
                        segments[segments.length - 1]![before] = { ...carrier, ruby: gloss };
                        if (base.elements('w:r').length > 1) {
                            this.report(
                                'unsupported-block',
                                'a w:ruby gloss over several runs was set over the first of them',
                            );
                        }
                    }
                    break;
                }
                case 'w:pict': {
                    // VML — the shape format DrawingML replaced, still written
                    // by older producers and by anything converting from .doc.
                    flush();
                    // A VML text box is a SHAPE with words in it, and its
                    // shape is usually a plain rectangle — so `readShape`
                    // below would happily draw the panel and drop the text.
                    const textBox = this.readTextBox(child, 'VML');
                    if (null !== textBox) {
                        this.pendingFloats.push(textBox);
                        break;
                    }
                    // An INLINE one sits in the line, so it is a piece rather
                    // than a float — and it carries its own stacked words, the
                    // way the DrawingML spelling does.
                    const inlineBox = this.readInlineVmlTextBox(child);
                    if (null !== inlineBox) {
                        segments[segments.length - 1]!.push({
                            ...resolved,
                            text: OBJECT_REPLACEMENT,
                            shape: inlineBox,
                        });
                        break;
                    }

                    // A box this cannot build — no width to stack against, or
                    // nothing but a table inside — still had words in it, and
                    // they are about to be dropped by the shape path below.
                    // The report belongs HERE rather than where it used to
                    // live: said unconditionally, it announced a loss on every
                    // box this now builds.
                    if (null !== findByLocalName(child, 'txbxContent')) {
                        this.report(
                            'unsupported-block',
                            'a VML text box held nothing this engine could lay out;'
                            + ' the box was kept and the text inside it dropped',
                        );
                    }
                    const picture = this.readPicture(child);
                    if (null !== picture) {
                        segments[segments.length - 1]!.push({
                            ...resolved,
                            text: OBJECT_REPLACEMENT,
                            image: picture,
                        });
                        break;
                    }

                    // No picture in it: the shape is FURNITURE — a rule, a
                    // border, a coloured panel — and a document whose boxes are
                    // VML had none of them before.
                    const box = this.readShape(child);
                    if (null !== box) {
                        segments[segments.length - 1]!.push({
                            ...resolved,
                            text: OBJECT_REPLACEMENT,
                            shape: box,
                        });
                    }
                    break;
                }
                case 'w:br':
                    if ('page' === child.attribute('w:type')) {
                        flush();
                        segments.push([]);
                    } else {
                        // A line break the author typed. The line breaker
                        // already treats this as a mandatory break.
                        text += '\n';
                    }
                    break;
                case 'w:softHyphen':
                    // An offer of a break, carried as the character itself so
                    // the breaker can take it or leave it. It has no
                    // width where it is left, and becomes a drawn hyphen where
                    // it is taken.
                    text += SOFT_HYPHEN_TEXT;
                    break;
                case 'w:cr':
                    // The same break under an older name, and it was in no
                    // branch at all — so the text either side of one ran
                    // together. Measured: `cc` and `dd` either side
                    // of a `w:cr` printed 11.50 apart, exactly as they do
                    // either side of a `w:br`.
                    text += '\n';
                    break;
                default:
                    break;
            }
        }
        flush();
    }

    /**
     * A `w:drawing`, if it holds a picture this engine can draw.
     *
     * Descends by LOCAL name. The prefixes are `wp`, `a` and `pic` by
     * convention and by convention only — a producer may bind those namespaces
     * to any prefix it likes, and matching on the prefix would silently lose
     * every picture in such a file.
     *
     * Only `wp:inline` is read. A `wp:anchor` is a FLOATING picture, which
     * needs text to wrap around it; placing one inline instead would push the
     * text it was meant to sit beside down the page.
     */
    /**
     * A `w:pict` — a picture in the shape language DrawingML replaced.
     *
     * The picture is named the same way a drawing names one, so it resolves
     * through the same relationship map. What differs is the SIZE: VML states
     * it in the shape's CSS-ish `style` rather than in EMU, and Word writes
     * points there where a converted document may say inches or centimetres.
     */
    private readPicture(pict: XmlElement): InlineImage | null {
        const data = findByLocalName(pict, 'imagedata');
        const id = data?.attribute('r:id') ?? null;
        const content = null === id ? undefined : this.mediaById[id];

        if (undefined === content) {
            if (null !== id) {
                this.report(
                    'unsupported-block',
                    `a VML picture referenced "${id}", which the package did not supply`,
                );
            }

            return null;
        }

        const shape = findByLocalName(pict, 'shape') ?? findByLocalName(pict, 'rect');
        const style = shape?.attribute('style') ?? '';
        const description = shape?.attribute('alt') ?? null;

        return {
            widthPx: cssLengthToPx(styleValue(style, 'width')) ?? 0,
            heightPx: cssLengthToPx(styleValue(style, 'height')) ?? 0,
            content: null === description ? content : { ...content, description },
        };
    }

    /**
     * A text box: the words inside a `w:txbxContent`, and the frame round them.
     *
     * Both spellings wrap that same element — VML's `v:textbox` and
     * DrawingML's `wps:txbx` — so ONE search finds either, and only the frame
     * around it differs. Words that cannot be placed are REPORTED rather than
     * dropped in silence, which is what this used to do.
     *
     * Measured against LibreOffice, whose numbers the placement reproduces: a
     * box anchored 200pt across and 20pt down printed its first baseline at
     * 279.30, 736.89 — the frame's corner, plus the inset, plus the line's own
     * rise — and the body either side stepped its ordinary 11.50, undisturbed.
     *
     * `null` means this drawing is not a text box at all, which is the ordinary
     * answer: every picture in every document comes through here.
     */
    private readTextBox(drawing: XmlElement, spelling: 'VML' | 'DrawingML'): FloatingBox | null {
        const content = findByLocalName(drawing, 'txbxContent');
        if (null === content) {
            return null;
        }

        const frame = 'VML' === spelling
            ? this.vmlBoxFrame(drawing)
            : this.drawingBoxFrame(drawing);

        if (null === frame) {
            // An INLINE box is not a float and is not a fault: it sits in the
            // line, and the caller reads it as one. Saying it was
            // dropped would be untrue, and saying nothing where it really IS
            // dropped would be worse, so the three cases are told apart here.
            if (null !== byLocalName(drawing, 'inline')) {
                return null;
            }

            // The VML spelling of the same thing — a shape with no
            // `position:absolute` — which this engine does not build YET.
            // Dropped, and said in its own words: the generic message below
            // would claim it said nothing about where it sits, and it did. It
            // said "in the line".
            //
            // MEASURED, and mis-read twice before it was.
            // The extra room such a box takes is a CONSTANT: printed at 30, 90
            // and 180pt and again beside a wrapping paragraph, it is the stated
            // width plus **18.0** every time, and the words sit **13.2** in
            // from the box's left edge.
            //
            // TWO decimals, not three. The prints give positions to a hundredth
            // and this engine already runs about 0.10 left of them, so 18.01
            // and 18.02 are the same measurement — the printed page gives
            // 18.02 and `run-width.mjs` says 18.01 off the wrap probe.
            // Agreement in that last digit would be arithmetic, not evidence.
            //
            // The 17.50 read first and the 16.40 read next were the same
            // mistake twice: the box's start was ESTIMATED from the width of
            // the text before it rather than computed from the face. `P-before`
            // is 34.43, not the 36 that was assumed, and `L-` is 9.44. The page
            // never moved; the arithmetic did. Measure the run you are
            // subtracting, or subtract nothing — `tools/probes/run-width.mjs`
            // exists so that subtraction is never estimated again.
            //
            // The HEIGHT was carried here as a missing number — "stepped 38.25
            // where a plain line steps 11.55" — and it never was one. The shape
            // falls through to the inline-shape path and this engine already
            // steps such a line by box-height-plus-descent, which is precisely
            // what LibreOffice does — measured at three sizes, in two fonts.
            //
            // Both are built now — the room and the words alike — so this
            // hands the inline spelling back for the caller to place
            // in the line, and says nothing, because nothing was lost.
            //
            if ('VML' === spelling) {
                return null;
            }

            // Named by SPELLING, and not only for the reader's sake: the two
            // travel different paths through this file, and one message for
            // both meant a test could not tell which of them had spoken —
            // every mutation silencing one of the two survived.
            this.report(
                'unsupported-block',
                `a ${spelling} text box said nothing about where it sits;`
                + ' it was dropped, with the text inside it',
            );

            return null;
        }

        const blocks = this.readBoxBlocks(content, spelling);

        return 0 === blocks.length ? null : { ...frame, blocks };
    }

    /**
     * The same box, sitting IN the line rather than beside it — `wp:inline`.
     *
     * Measured: a 90x36pt box takes exactly its `wp:extent` in the
     * line and draws its words 7.20 inside that, which is the body inset. So
     * it is a SHAPE with words: the piece it becomes already measures its own
     * width and grows its line, and only the words needed somewhere to live.
     *
     * They are stacked HERE, unlike a float's, and can be: an inline box's
     * inner width is settled by its extent before anything knows where the box
     * lands, where a float's content is stacked at placement in `floatContent`.
     */
    private readInlineTextBox(drawing: XmlElement): InlineShape | null {
        const content = findByLocalName(drawing, 'txbxContent');
        const inline = byLocalName(drawing, 'inline');
        if (null === content || null === inline) {
            return null;
        }

        const extent = byLocalName(inline, 'extent');
        const widthPx = emuToPx(attributeNumber(extent, 'cx') ?? 0);
        const heightPx = emuToPx(attributeNumber(extent, 'cy') ?? 0);
        const blocks = this.readBoxBlocks(content, 'DrawingML');
        if (0 === blocks.length || widthPx <= 0) {
            return null;
        }

        const inset = readBodyInset(findByLocalName(inline, 'bodyPr'));
        const stacked = stackBlocks(
            blocks,
            Math.max(0, widthPx - inset.leftPx - inset.rightPx),
        );

        return {
            widthPx,
            heightPx,
            strokeWidthPx: 0,
            lines: stacked.lines.map((line) => translateLine(line, inset.leftPx, inset.topPx)),
            rows: stacked.rows.map((row) => translateRow(row, inset.leftPx, inset.topPx)),
        };
    }

    /**
     * The VML spelling of an inline text box: a `v:shape` with words in it and
     * no `position:absolute`.
     *
     * ## Every number here was printed, and the last of them was aimed at
     *
     * The box keeps 9pt of wrap distance either side of itself and
     * an INSET of 4.25 on every side. That inset was arrived at
     * twice over, from measurements that knew nothing of each other: the words
     * start 13.2 from the shape's origin — 13.21, 13.21 and 13.22 at 90pt@10pt,
     * 150pt@10pt and 90pt@20pt, so neither the font's nor the box's — which is
     * the 9.0 of wrap plus 4.25; and the first line sits 13.55 below the box's
     * top at 10pt and 22.90 at 20pt, which are the same 4.25 once the line's
     * own rise comes off.
     *
     * The content width is the stated width less that inset twice. Two rules
     * survived the whole arc — `stated - 13.2` and `stated - 8.5` — because
     * every wrap measured fell outside the 4.7pt gap between them. A string
     * built with `run-width.mjs` to land INSIDE it settled it: `alpha beta
     * gammaw`, 79.97 wide, stayed on its line in a 90pt box, which `stated -
     * 13.2` (76.8) forbids.
     *
     * `null` for a float — that is `readTextBox`'s — and for a shape holding
     * no words, which is furniture and `readShape`'s.
     */
    private readInlineVmlTextBox(pict: XmlElement): InlineShape | null {
        const content = findByLocalName(pict, 'txbxContent');
        const shape = findByLocalName(pict, 'shape')
            ?? findByLocalName(pict, 'rect')
            ?? findByLocalName(pict, 'roundrect');
        if (null === content || null === shape) {
            return null;
        }

        const style = shape.attribute('style') ?? '';
        if ('absolute' === styleValue(style, 'position')) {
            return null;
        }

        const widthPx = cssLengthToPx(styleValue(style, 'width'));
        const heightPx = cssLengthToPx(styleValue(style, 'height'));
        if (null === widthPx || null === heightPx || widthPx <= 0) {
            return null;
        }

        const blocks = this.readBoxBlocks(content, 'VML');
        if (0 === blocks.length) {
            return null;
        }

        const insetPx = 4.25 * 96 / 72;
        const stacked = stackBlocks(blocks, Math.max(0, widthPx - insetPx - insetPx));

        // A box is a shape that happens to hold words, so its PANEL is read
        // the way any shape's is — including the rule that a `v:shape` naming
        // no shapetype paints nothing. Reading it here instead would
        // give a filled `v:rect` with text in it a box nobody draws, which the
        // print does draw: the green one in `vml-shape-geometryless.docx`.
        const painted = this.readShape(pict);
        const leadPx = cssLengthToPx(styleValue(style, 'mso-wrap-distance-left')) ?? 9 * 96 / 72;

        return {
            ...(painted ?? { widthPx, heightPx, strokeWidthPx: 0 }),
            advanceWidthPx: widthPx + this.vmlWrapExtraPx(style),
            leadPx,
            lines: stacked.lines.map((line) => translateLine(line, leadPx + insetPx, insetPx)),
            rows: stacked.rows.map((row) => translateRow(row, leadPx + insetPx, insetPx)),
        };
    }

    /**
     * The room a VML shape keeps in the line BESIDE the width it draws at.
     *
     * Wrap distance, 9pt a side where the document says nothing — which is
     * what a document usually says, and why this read as an unexplained
     * constant for five slices.
     */
    private vmlWrapExtraPx(style: string): number {
        const defaultPx = 9 * 96 / 72;

        return (cssLengthToPx(styleValue(style, 'mso-wrap-distance-left')) ?? defaultPx)
            + (cssLengthToPx(styleValue(style, 'mso-wrap-distance-right')) ?? defaultPx);
    }

    /**
     * What is inside the box, read as an ordinary run of blocks.
     *
     * ## The outer paragraph's floats have to survive this
     *
     * `readParagraph` clears the reader's pending floats when it starts and
     * takes them when it ends — and the box is found part-way through the OUTER
     * paragraph's runs, so reading the paragraphs inside it would carry off
     * whatever that paragraph had anchored already. Saved and put back, or a
     * picture anchored before a text box in the same paragraph disappears.
     *
     * A table inside a box is dropped and said so: `PlacedFloat` carries lines,
     * and rows would want the renderer's row path as well as its line path.
     */
    private readBoxBlocks(
        content: XmlElement,
        spelling: 'VML' | 'DrawingML',
    ): Block[] {
        const outer = this.pendingFloats;
        const blocks: Block[] = [];

        for (const child of blockChildren(content)) {
            if ('w:p' === child.name) {
                blocks.push(...this.readParagraph(child));
            } else if ('w:tbl' === child.name) {
                blocks.push(this.readTable(child));
            }
        }

        this.pendingFloats = outer;

        return blocks;
    }

    /**
     * A DrawingML text box's frame: `wp:anchor` for where, `wp:extent` for how
     * big, `wps:bodyPr` for the space inside it.
     *
     * An INLINE text box — `wp:inline` rather than `wp:anchor` — sits in the
     * line like a picture rather than beside the text. It returns null and the
     * caller says so.
     *
     * MEASURED at last, where this said no measurement covered it.
     * A 90x36pt inline box takes **exactly its `wp:extent`** in the line: the
     * print draws its words at 115.35 — the box's own left edge plus the 7.20
     * body inset — and the text after it at **198.20**, which is 90.00 past
     * where the box begins. This engine reserves nothing and puts that text at
     * 108.09, so a box's worth of words vanishes and the line closes over it.
     *
     * NOT built here, and what it needs is now known rather than guessed: a
     * line piece that carries BLOCKS, laid out at layout time as
     * `floatContent` already lays a float's, rather than the size-and-no-glyphs
     * piece a shape gets. The drop is reported meanwhile, which is the least a
     * reader can do about words it will not draw.
     *
     * The VML spelling wants its own answer: the same 90pt shape reserved
     * 108.0 and drew its words 13.2 in, and nothing in the file accounts for
     * the extra 18.0 — read as 17.50 and then 16.40 before it was measured
     * rather than estimated.
     */
    private drawingBoxFrame(drawing: XmlElement): Omit<FloatingBox, 'blocks'> | null {
        const anchor = byLocalName(drawing, 'anchor');
        if (null === anchor) {
            return null;
        }

        const extent = byLocalName(anchor, 'extent');

        return {
            widthPx: emuToPx(attributeNumber(extent, 'cx') ?? 0),
            heightPx: emuToPx(attributeNumber(extent, 'cy') ?? 0),
            inset: readBodyInset(findByLocalName(anchor, 'bodyPr')),
            horizontal: readFloatPosition(byLocalName(anchor, 'positionH')),
            vertical: readFloatPosition(byLocalName(anchor, 'positionV')),
            wrap: readWrap(anchor),
            marginTopPx: emuToPx(attributeNumber(anchor, 'distT') ?? 0),
            marginBottomPx: emuToPx(attributeNumber(anchor, 'distB') ?? 0),
            marginLeftPx: emuToPx(attributeNumber(anchor, 'distL') ?? 0),
            marginRightPx: emuToPx(attributeNumber(anchor, 'distR') ?? 0),
            behindText: '1' === (anchor.attribute('behindDoc') ?? '0'),
        };
    }

    /**
     * A VML text box's frame, which says everything in CSS.
     *
     * `margin-left` and `margin-top` are the offsets and they are measured from
     * the defaults this engine already uses — the column across, the paragraph
     * down. Measured: `margin-left:200pt` on a page with a 72pt
     * margin put the frame's edge at 272.10, which is the column's own left
     * plus 200.
     *
     * A shape may state some OTHER origin in `mso-position-horizontal-relative`
     * and its vertical twin. Those are not measured, so they are not obeyed —
     * and the box says so rather than landing somewhere unexplained.
     */
    private vmlBoxFrame(pict: XmlElement): Omit<FloatingBox, 'blocks'> | null {
        const shape = findByLocalName(pict, 'shape')
            ?? findByLocalName(pict, 'rect')
            ?? findByLocalName(pict, 'roundrect');
        if (null === shape) {
            return null;
        }

        const style = shape.attribute('style') ?? '';
        const widthPx = cssLengthToPx(styleValue(style, 'width'));
        const heightPx = cssLengthToPx(styleValue(style, 'height'));
        if (null === widthPx || null === heightPx) {
            return null;
        }

        // A VML shape with no absolute position sits IN the line, and this
        // engine builds that only for the DrawingML spelling. It was
        // coming through here as a float with no offsets at all — placed at
        // the paragraph's own origin, and its words drawn nowhere. Dropped and
        // SAID, which is what the other spelling did before it was built.
        //
        // Returning null drops it from the FLOAT path only. It then falls
        // through to the ordinary inline-shape path and reserves its stated
        // size — which is why its LINE is already right: this engine steps an
        // inline shape by box-height-plus-descent, and LibreOffice was
        // measured doing exactly that, at three sizes and in two fonts.
        //
        // What is still missing is the 18.0 of extra width the print gives
        // such a box, and the words inside it — an extra read as 17.50 and
        // then 16.40 before it was measured rather than estimated.
        if ('absolute' !== styleValue(style, 'position')) {
            return null;
        }

        if (null !== styleValue(style, 'mso-position-horizontal-relative')
            || null !== styleValue(style, 'mso-position-vertical-relative')) {
            this.report(
                'unsupported-block',
                'a VML text box named an origin of its own; it was placed against the column'
                + ' and the paragraph instead',
            );
        }

        return {
            widthPx,
            heightPx,
            inset: readVmlInset(findByLocalName(pict, 'textbox')?.attribute('inset') ?? null),
            horizontal: {
                relativeTo: 'column',
                offsetPx: cssLengthToPx(styleValue(style, 'margin-left')) ?? 0,
            },
            vertical: {
                relativeTo: 'paragraph',
                offsetPx: cssLengthToPx(styleValue(style, 'margin-top')) ?? 0,
            },
            // A VML float wraps only where a `w10:wrap` says so. Measured: with
            // none of them, the body stepped its ordinary 11.50 past both boxes.
            wrap: null === findByLocalName(pict, 'wrap') ? 'none' : 'square',
            marginTopPx: 0,
            marginBottomPx: 0,
            marginLeftPx: 0,
            marginRightPx: 0,
            behindText: false,
        };
    }

    /**
     * A `w:pict` that holds no picture: a shape DRAWN rather than shown.
     *
     * VML says the size in the shape's `style` and its colours in attributes
     * beside it — `fillcolor`, `strokecolor`, `strokeweight` — and says so in
     * CSS spellings rather than in OOXML's. A shape with neither fill nor
     * stroke is invisible, and is left out rather than given a box of nothing.
     */
    private readShape(pict: XmlElement): InlineShape | null {
        const shape = findByLocalName(pict, 'shape')
            ?? findByLocalName(pict, 'rect')
            ?? findByLocalName(pict, 'roundrect');
        if (null === shape) {
            return null;
        }

        const style = shape.attribute('style') ?? '';
        const widthPx = cssLengthToPx(styleValue(style, 'width'));
        const heightPx = cssLengthToPx(styleValue(style, 'height'));
        if (null === widthPx || null === heightPx) {
            return null;
        }

        // A `v:shape` has no outline of its OWN: it takes one from the
        // `v:shapetype` its `type` names. Name no type and there is no path to
        // draw, so LibreOffice paints nothing at all — while still keeping the
        // shape's room in the line.
        //
        // This engine painted a white box with a black border there, because
        // the colours below default to white and black for a shape that names
        // none. Measured with the control that makes the null mean something:
        // in ONE print, a `v:rect` — which IS a rectangle and needs no
        // shapetype — came out with its 3pt stroke at width 3.000, and the
        // bare `v:shape` beside it produced no mark anywhere on the page.
        //
        // The room is kept deliberately: the print puts the run after such a
        // shape at the same 189.55 whether anything is drawn or not, and that
        // room is what was measured. A TYPED shape is untouched by this.
        // An inline VML shape keeps WRAP DISTANCE either side of itself, and
        // states the width it is DRAWN at. Those are two numbers, so widening
        // the drawn one would grow a box the print draws at its stated size.
        //
        // The 18.0 carried for a long time as an unexplained constant is this,
        // defaulted: 9pt a side. ISOLATED at last — a 90pt rect
        // whose run after it starts at 189.55 by default starts at 171.55 with
        // both distances set to 0, at 211.55 with both at 20pt, and at 191.55
        // with 20pt on the left alone. The 20pt line is the control that makes
        // the 0 line mean something, which the earlier inset probe never got.
        //
        // It also settles a contradiction the suite caught: a shape in a
        // TURNED cell was measured charging its height and nothing more. Wrap
        // distance is per-SIDE, and a turned line advances down the page, where
        // the top and bottom distances apply — and those default to nought.
        // Hence `advanceWidthPx`, which the turned path does not read.
        //
        // Where the box sits INSIDE that room is not built, because it is not
        // measured: the print puts a 90pt box's edge about 8.8 in, but the only
        // fixtures that draw at all draw degenerately — one half-width segment
        // — so this leaves the box against the left of its room, where it was.
        const advanceWidthPx = widthPx + this.vmlWrapExtraPx(style);

        const generic = findByLocalName(pict, 'shape');
        if (null !== generic && generic === shape && null === generic.attribute('type')) {
            return { widthPx, heightPx, strokeWidthPx: 0, advanceWidthPx };
        }

        // `filled`/`stroked` are how VML says NO, and the colours default to
        // white and black where the shape says nothing at all.
        const fillHex = 'f' === shape.attribute('filled')
            ? null
            : normaliseVmlColour(shape.attribute('fillcolor') ?? '#FFFFFF');
        const strokeHex = 'f' === shape.attribute('stroked')
            ? null
            : normaliseVmlColour(shape.attribute('strokecolor') ?? '#000000');
        const strokeWidthPx = cssLengthToPx(shape.attribute('strokeweight')) ?? 1;

        if (null === fillHex && null === strokeHex) {
            return null;
        }

        return {
            widthPx,
            heightPx,
            ...(null === fillHex ? {} : { fillHex }),
            ...(null === strokeHex ? {} : { strokeHex }),
            strokeWidthPx,
            advanceWidthPx,
        };
    }

    private readDrawing(drawing: XmlElement): InlineImage | null {
        const inline = byLocalName(drawing, 'inline');
        if (null === inline) {
            return null;
        }

        const extent = byLocalName(inline, 'extent');
        const blip = descendByLocalName(inline, ['graphic', 'graphicData', 'pic', 'blipFill', 'blip']);
        const id = blip?.attribute('r:embed') ?? null;
        const content = null === id ? undefined : this.mediaById[id];

        if (undefined === content) {
            this.report(
                'unsupported-block',
                null === id
                    ? 'a drawing had no embedded picture'
                    : `a picture referenced "${id}", which the package did not supply`,
            );

            return null;
        }

        const description = byLocalName(inline, 'docPr')?.attribute('descr') ?? null;

        return {
            // The size the DOCUMENT asks for, not the picture's own: a large
            // photograph scaled into a small box is the size of the box.
            widthPx: emuToPx(attributeNumber(extent, 'cx') ?? 0),
            heightPx: emuToPx(attributeNumber(extent, 'cy') ?? 0),
            content: null === description ? content : { ...content, description },
        };
    }

    /**
     * A `wp:anchor` — a picture text flows around rather than through.
     *
     * Returns null for a drawing that is not anchored, so the caller can fall
     * through to the inline reader without asking twice.
     */
    private readAnchor(drawing: XmlElement): FloatingImage | null {
        const anchor = byLocalName(drawing, 'anchor');
        if (null === anchor) {
            return null;
        }

        const extent = byLocalName(anchor, 'extent');
        const blip = descendByLocalName(anchor, ['graphic', 'graphicData', 'pic', 'blipFill', 'blip']);
        const id = blip?.attribute('r:embed') ?? null;
        const content = null === id ? undefined : this.mediaById[id];

        if (undefined === content) {
            this.report('unsupported-block', 'a floating picture had no embedded image');

            return null;
        }

        const description = byLocalName(anchor, 'docPr')?.attribute('descr') ?? null;

        return {
            image: {
                widthPx: emuToPx(attributeNumber(extent, 'cx') ?? 0),
                heightPx: emuToPx(attributeNumber(extent, 'cy') ?? 0),
                content: null === description ? content : { ...content, description },
            },
            horizontal: readFloatPosition(byLocalName(anchor, 'positionH')),
            vertical: readFloatPosition(byLocalName(anchor, 'positionV')),
            wrap: readWrap(anchor),
            marginTopPx: emuToPx(attributeNumber(anchor, 'distT') ?? 0),
            marginBottomPx: emuToPx(attributeNumber(anchor, 'distB') ?? 0),
            marginLeftPx: emuToPx(attributeNumber(anchor, 'distL') ?? 0),
            marginRightPx: emuToPx(attributeNumber(anchor, 'distR') ?? 0),
            behindText: '1' === (anchor.attribute('behindDoc') ?? '0'),
        };
    }

    /**
     * The `w:rt` of a `w:ruby` — the gloss itself.
     *
     * Its runs are resolved exactly as any other run's are, so the gloss is in
     * whatever face and size the file gave it. `w:rubyPr/w:hps` states that
     * size a second time and is used only where the run itself is silent:
     * measured against LibreOffice, a gloss whose run says 5pt is
     * drawn at 5pt, which is what both said in the file that was probed.
     */
    private readRubyGloss(ruby: XmlElement, styleId: string | null): RubyGloss | undefined {
        const rt = byLocalName(ruby, 'rt');
        if (null === rt) {
            return undefined;
        }

        let text = '';
        for (const run of rt.elements('w:r')) {
            for (const child of run.elements()) {
                if ('w:t' === child.name) {
                    text += child.text;
                }
            }
        }

        if ('' === text) {
            return undefined;
        }

        const first = rt.elements('w:r')[0] ?? null;
        const characterStyle = first?.element('w:rPr')?.element('w:rStyle')?.attribute('w:val') ?? null;
        const stated = attributeNumber(descendByLocalName(ruby, ['rubyPr', 'hps']), 'w:val');
        const properties: RunProperties = {
            ...this.styles.runProperties(styleId, characterStyle, this.tableStyleLends.run),
            ...readDirectRunProperties(first?.element('w:rPr') ?? null),
        };
        const resolved = this.resolveRun({
            ...properties,
            ...(undefined === properties.sizeHalfPoints && null !== stated
                ? { sizeHalfPoints: stated }
                : {}),
        });

        return {
            text,
            font: resolved.font,
            sizePx: resolved.sizePx,
            ...(undefined === resolved.colorHex ? {} : { colorHex: resolved.colorHex }),
        };
    }

    private resolveRun(properties: RunProperties): Omit<StyledRun, 'text'> {
        const resolved = this.fonts.resolve(
            properties.family ?? null,
            true === properties.bold,
            true === properties.italic,
        );

        if (null !== resolved.substitutedFor) {
            this.report(
                'font-substituted',
                `"${resolved.substitutedFor}" is not available; using the metric-compatible ${resolved.family}`,
            );
        }

        const nominalPx = undefined === properties.sizeHalfPoints
            ? halfPointsToPx(this.fonts.defaultSizePt * 2)
            : halfPointsToPx(properties.sizeHalfPoints);

        // A script run is SMALLER and sits off the line. Reducing the size here
        // rather than at drawing time means every measurement downstream —
        // widths, line breaking, the line's own height — already knows.
        // `w:position` moves a run by hand, and WINS over `w:vertAlign`
        // outright — measured against LibreOffice, a run that stated both came
        // out at the hand-set height and at FULL size, its neighbour starting
        // exactly where an unscripted one did. The two do not add.
        const positioned = undefined !== properties.positionHalfPoints
            && 0 !== properties.positionHalfPoints;
        const script = !positioned
            && ('superscript' === properties.vertAlign || 'subscript' === properties.vertAlign);
        const sizePx = script ? nominalPx * SCRIPT_SCALE : nominalPx;

        // Half-points of the LINE: twelve units raised a run 6.90pt in a font
        // whose line stands 1.15 times its size, and 7.00 in one at 1.167.
        const shiftPx = positioned
            ? -halfPointsToPx(properties.positionHalfPoints ?? 0)
                * (resolved.font.naturalLineHeight(nominalPx) / nominalPx)
            : ('superscript' === properties.vertAlign
                ? -nominalPx * SUPERSCRIPT_RISE
                : ('subscript' === properties.vertAlign ? nominalPx * SUBSCRIPT_DROP : 0));

        return {
            font: resolved.font,
            sizePx,
            ...(0 === shiftPx ? {} : { baselineShiftPx: shiftPx }),
            ...(undefined === properties.colorHex ? {} : { colorHex: properties.colorHex }),
            ...(undefined === properties.highlightHex
                ? {}
                : { highlightHex: properties.highlightHex }),
            ...(undefined === properties.border ? {} : { border: properties.border }),
            ...(undefined === properties.underline ? {} : { underline: properties.underline }),
            ...(undefined === properties.strike ? {} : { strike: properties.strike }),
            // `w:spacing` on a run is TWENTIETHS of a point, which is the same
            // unit `w:sz` uses halves of and nothing else in the file uses.
            ...(undefined === properties.letterSpacingTwentieths
                || 0 === properties.letterSpacingTwentieths
                ? {}
                : { letterSpacingPx: pointsToPx(properties.letterSpacingTwentieths / 20) }),
            ...(true === properties.kerned ? { kerned: true } : {}),
        };
    }

    private paragraphStyle(
        properties: ParagraphProperties,
        runs: readonly StyledRun[],
        position: { first: boolean; last: boolean },
    ): ParagraphStyle {
        const lineHeightPx = resolveLineHeight(properties, runs);

        return {
            // Space before belongs to the START of the paragraph and space
            // after to its END. A paragraph split by a page break must not
            // repeat either at the seam.
            spaceBeforePx: position.first ? twipsToPx(properties.spaceBeforeTwips ?? 0) : 0,
            spaceAfterPx: position.last ? twipsToPx(properties.spaceAfterTwips ?? 0) : 0,
            indentLeftPx: twipsToPx(properties.indentLeftTwips ?? 0),
            indentRightPx: twipsToPx(properties.indentRightTwips ?? 0),
            indentFirstLinePx: position.first ? twipsToPx(properties.indentFirstLineTwips ?? 0) : 0,
            pageBreakBefore: position.first ? true === properties.pageBreakBefore : true,
            keepLinesTogether: true === properties.keepLines,
            // Only the LAST piece of a paragraph split across a page can be
            // kept with what follows it; the earlier pieces are followed by the
            // rest of their own paragraph.
            keepWithNext: position.last && true === properties.keepNext,
            ...(true === properties.suppressLineNumbers ? { suppressLineNumbers: true } : {}),
            // Only the FIRST piece of a paragraph split across a page opens the
            // box; what carries over is inside one already.
            ...(undefined === properties.borders ? {} : { borders: properties.borders }),
            ...(undefined === properties.tabStops
                ? {}
                : {
                    tabStops: properties.tabStops.map((stop) => ({
                        positionPx: twipsToPx(stop.positionTwips),
                        align: stop.align,
                        ...(undefined === stop.leader ? {} : { leader: stop.leader }),
                    })),
                }),
            ...(null === this.decimalSymbol ? {} : { decimalSymbol: this.decimalSymbol }),
            defaultTabPx: this.defaultTabPx,
            ...(null === lineHeightPx ? {} : { lineHeightPx }),
            ...(undefined === properties.lineRule ? {} : { lineRule: properties.lineRule }),
            ...(undefined === properties.widowControl ? {} : { widowControl: properties.widowControl }),
            ...(undefined === properties.alignment ? {} : { alignment: properties.alignment }),
        };
    }

    report(kind: DiagnosticKind, detail: string): void {
        // A 160-paragraph document would otherwise report the same missing font
        // 160 times and bury everything else.
        const key = `${kind}:${detail}`;
        if (this.reported.has(key)) {
            return;
        }
        this.reported.add(key);
        this.diagnostics.push({ kind, detail });
    }
}

/**
 * Turn `w:spacing` into one exact line height.
 *
 * Word has three modes and they are not interchangeable. `auto` is a MULTIPLE
 * of single spacing expressed in 240ths, so 360 is 1.5 lines; `exact` is a
 * fixed height that clips tall text; `atLeast` is a floor. Resolving all three
 * to a number here keeps the layout model to a single concept.
 */
function resolveLineHeight(properties: ParagraphProperties, runs: readonly StyledRun[]): number | null {
    const line = properties.lineTwips;
    if (undefined === line) {
        return null;
    }

    const natural = naturalLineHeightOf(runs);

    // Word treats a missing lineRule as "auto".
    switch (properties.lineRule ?? 'auto') {
        case 'exact':
            return twipsToPx(line);
        case 'atLeast':
            return Math.max(natural, twipsToPx(line));
        default:
            return (natural * line) / 240;
    }
}

/**
 * Single spacing for these runs: the tallest natural line height among them.
 *
 * Shared by the two places that need it — resolving a paragraph's own
 * `w:spacing`, and re-resolving that same spacing against a section's grid —
 * so the second cannot drift from the first.
 */
function naturalLineHeightOf(runs: readonly StyledRun[]): number {
    let natural = 0;
    for (const run of runs) {
        natural = Math.max(natural, run.font.naturalLineHeight(run.sizePx));
    }

    return natural;
}

/**
 * U+00AD, spelled out rather than typed.
 *
 * A soft hyphen is INVISIBLE in a source file, and a reader meeting one in
 * a string literal cannot tell it from nothing at all.
 */
const SOFT_HYPHEN_TEXT = String.fromCodePoint(SOFT_HYPHEN);

/**
 * Word's default gap between columns, for a `w:cols` that states none.
 *
 * Half an inch, and 1440 twips to the inch as OOXML defines — so this is a
 * DEFINITION rather than anything read off a page. (It leant on a citation in
 * a neighbouring comment until a constant was inserted between the two and
 * `audit:claims` noticed the justification had drifted out of view.)
 */
const DEFAULT_COLUMN_GAP_TWIPS = 720;

/**
 * `w:cols/@w:equalWidth`, read as the toggle it is.
 *
 * ABSENT is false here, which is the opposite of what the schema says the
 * default is — see {@link readColumns}. It is what the page does, and the
 * default only decides the case where no width is stated, which this is never
 * asked about.
 */
function isEqualWidth(value: string | null): boolean {
    return null !== value && !('0' === value || 'false' === value || 'off' === value);
}

/**
 * `w:cols` — the columns a section's text flows down.
 *
 * Returned as absolute edges rather than a count and a gap, because Word can
 * state each column's own width and a count could not say that. Equal columns
 * divide what is left of the writing width once the gaps are taken out of it,
 * which is the arithmetic LibreOffice printed at two, three and a wide gap.
 */
function readColumns(
    sectPr: XmlElement | null,
    writingTwips: number,
): PageColumn[] | null {
    const cols = sectPr?.element('w:cols') ?? null;
    if (null === cols) {
        return null;
    }

    const stated = cols.elements('w:col');
    const count = attributeNumber(cols, 'w:num') ?? (stated.length || 1);
    if (count < 2 && 0 === stated.length) {
        return null;
    }

    const columns: PageColumn[] = [];
    let leftTwips = 0;

    // A stated width WINS unless the document explicitly asks for equal
    // columns — which is not what the schema's default says, and is what the
    // page does. The same two columns of 3000 and 5526 twips put
    // their second at 247.10 with `w:equalWidth="0"`, with `"false"`, and with
    // the attribute absent altogether; only `"1"` divided the width evenly and
    // put it at 315.75.
    //
    // Reading it as "off" alone missed two of those three: the attribute is a
    // TOGGLE, and a toggle is spelt `0`, `false` or `off` — while a document
    // that states widths and says nothing at all plainly means the widths.
    if (stated.length > 0 && !isEqualWidth(cols.attribute('w:equalWidth'))) {
        for (const col of stated) {
            const widthTwips = attributeNumber(col, 'w:w') ?? 0;
            columns.push({ leftPx: twipsToPx(leftTwips), widthPx: twipsToPx(widthTwips) });
            leftTwips += widthTwips + (attributeNumber(col, 'w:space') ?? 0);
        }

        return columns;
    }

    const gapTwips = attributeNumber(cols, 'w:space') ?? DEFAULT_COLUMN_GAP_TWIPS;
    const widthTwips = (writingTwips - gapTwips * (count - 1)) / count;
    for (let index = 0; index < count; index++) {
        columns.push({
            leftPx: twipsToPx(index * (widthTwips + gapTwips)),
            widthPx: twipsToPx(widthTwips),
        });
    }

    return columns;
}

/**
 * `w:pgNumType/@w:start`, when a section restarts its numbering.
 *
 * The element is often there and EMPTY — the fixture Word saved for this
 * repository carries one — which says nothing at all and must not be read as
 * a restart at zero.
 */
/**
 * The gap LibreOffice leaves when `w:distance` says nothing: half a centimetre.
 *
 * Measured, not assumed — a number with no distance stated came out 14.05pt in
 * from the writing area, against the 18 a quarter-inch would have given.
 */
const DEFAULT_LINE_NUMBER_GAP_PX = 0.5 / 2.54 * 96;

function readLineNumbering(sectPr: XmlElement | null): LineNumbering | null {
    const element = sectPr?.element('w:lnNumType') ?? null;
    if (null === element) {
        return null;
    }

    const restart = element.attribute('w:restart');
    const distance = attributeNumber(element, 'w:distance');
    const countBy = attributeNumber(element, 'w:countBy');

    return {
        // A `countBy` of nought would divide by zero, and says nothing anyway.
        countBy: null === countBy || countBy < 1 ? 1 : countBy,
        // `w:start` is ADDED to the count rather than being the first number:
        // stated as five with a `countBy` of two, LibreOffice printed 6, 8 and
        // 10 against lines one, three and five.
        start: attributeNumber(element, 'w:start') ?? 0,
        distancePx: null === distance ? DEFAULT_LINE_NUMBER_GAP_PX : twipsToPx(distance),
        // `newPage` is the default, and was measured as one: with nothing said,
        // the second page began again at 1 where `continuous` ran on to 69.
        restart: 'continuous' === restart || 'newSection' === restart ? restart : 'newPage',
    };
}

/**
 * `w:docGrid` — an East Asian typesetting grid, which moves LATIN text too.
 *
 * Measured against LibreOffice: a 360-twip pitch stepped 18pt where the same
 * text ungridded steps 11.5, and `linesAndChars` did the same as `lines` for
 * Latin, its `w:charSpace` changing nothing. `default` leaves the text alone.
 *
 * Applied to the section's paragraphs rather than carried beside them: the
 * grid is what a paragraph falls back to, and one that states its own
 * `w:spacing` keeps it — so this is a default, and defaults belong on the
 * thing they default.
 */
function readLinePitchPx(sectPr: XmlElement | null): number | null {
    const grid = sectPr?.element('w:docGrid') ?? null;
    const type = grid?.attribute('w:type') ?? 'default';
    const pitch = attributeNumber(grid, 'w:linePitch');

    return null === grid || null === pitch || pitch <= 0
        || ('lines' !== type && 'linesAndChars' !== type)
        ? null
        : twipsToPx(pitch);
}

/**
 * A section's grid, applied to its paragraphs.
 *
 * ## The pitch is the section's SINGLE SPACING, not a default
 *
 * A paragraph that states nothing takes the pitch, which is what this did from
 * the start. A paragraph that states spacing of its OWN is re-resolved against
 * the pitch rather than left alone — measured against LibreOffice on
 * an 18pt grid over a font whose natural line is 11.50:
 *
 *   `atLeast` 12pt   stepped 18.00, which is max(12, PITCH)
 *   `auto` 1.5       stepped 27.00, which is 1.5 x PITCH — not 1.5 x 11.50
 *   `exact` 14pt     stepped 14.00, the one rule the grid does not touch
 *
 * Measuring the `exact` case ALONE reads as "a paragraph stating its own
 * spacing wins". It wins on HEIGHT, and only exact does; the other two
 * take the grid's unit and ask their own question of it. Reading it the other
 * way cost a third of the lines on a page under `atLeast`.
 *
 * `exact` keeps its rule as well as its height, because clipping is the whole
 * of what `exact` means. It still carries the pitch, which changes where its
 * baseline sits inside that height (15.62 down a 24pt line, where the flat
 * 0.8 ratio this engine uses off a grid gives 19.20).
 */
/**
 * A cell's paragraphs, with any `atLeast` FLOOR taken off them.
 *
 * Under a grid, a cell takes the font's own line and neither the pitch nor its
 * own floor — measured at floors of 8, 12, 20 and 24 over a natural line of
 * 11.55, all four printing 11.55. Only `atLeast` goes: a cell asking
 * for one and a half lines still printed 17.25, one and a half of the font, so
 * `auto` is honoured and left alone.
 *
 * Taking the rule AND the height off leaves the paragraph as one that never
 * asked for spacing at all, which is the natural line by construction.
 */
function withoutCellFloors(table: Table): Table {
    return {
        ...table,
        rows: table.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
                ...cell,
                paragraphs: cell.paragraphs.map((block) => {
                    if (isTable(block)) {
                        // A table nested in a cell is still in a cell.
                        return withoutCellFloors(block);
                    }
                    if ('atLeast' !== block.style.lineRule) {
                        return block;
                    }

                    const { lineRule: _rule, lineHeightPx: _height, ...style } = block.style;

                    return { ...block, style };
                }),
            })),
        })),
    };
}

function onGrid(
    blocks: readonly Block[],
    pitchPx: number,
    snaps: (block: Block) => boolean,
    modernDefaults: boolean,
): Block[] {
    return blocks.map((block) => {
        // A table's own paragraphs live inside its cells and are left to the
        // cell's own stacking; the grid is the SECTION's line, and a cell is
        // not on it. Measured at last rather than asserted: under an
        // 18pt grid a cell's paragraphs printed 11.50 apart, the font's own
        // line, while the body paragraph above the table stepped 18.00 — and a
        // cell asking for 1.5 lines took 17.25, one and a half of the FONT,
        // where the same request on the body takes one and a half pitches.
        //  AND YET a cell sees the pitch for `atLeast`: under the
        // same 18.00 grid, a cell paragraph stating a floor of 12.00 printed
        // an 18.00 step — the section's pitch — where this engine gives it the
        // 12.00 it asked for. The two prints do not contradict each other so
        // much as split the question: a cell is off the grid for the STEP it
        // takes (11.50, and 17.25 for one and a half) and on it for the FLOOR
        // a paragraph may not fall below.
        //
        //  TWO PRINTS DISAGREE, and neither is wrong. Asked at
        // four floors and two pitches, a cell's `atLeast` step looked exactly
        // like `max(floor, PITCH)`: floors of 8 and 12 printed 18.00 under an
        // 18.00 grid, 20 and 24 printed themselves, and under a 24.00 grid
        // both 12 and 20 printed 24.00 — tracking the pitch, not a number that
        // happened to be 18.
        //
        // Building that broke `doc-grid-cell.docx`, which has the SAME 18.00
        // pitch and the SAME 12.00 floor and prints 11.50 — LibreOffice
        // dropping the floor BELOW what the paragraph asked, which stood
        // unexplained. The one thing that differs between the two
        // files is the font: a natural line of 11.50 there against 10.35 here.
        //
        // So the rule is not `max(floor, pitch)`, and what it is has to
        // account for a print that goes UNDER the floor. A cell stays off the
        // grid until it does.
        //  SETTLED, and the two prints never disagreed. Asked at
        // four floors, in the body AND in a cell, with the grid and without
        // it: no grid honours the floor everywhere, cells included; a grid
        // gives the BODY `max(floor, pitch)` and gives a CELL the font's own
        // line, ignoring the pitch and the floor both. The `max(floor, pitch)`
        // measured earlier is a BODY rule that was read off a cell.
        //
        // It was read off a cell in `grid-cell-floor.docx`, which is
        // hand-built and carries no `word/settings.xml` — and that part's mere
        // PRESENCE decides which defaults LibreOffice uses — an empty one
        // behaves like a full one. Only the modern branch is built: it is
        // what Word and LibreOffice both emit, and the legacy one is the
        // behaviour of a file no producer writes.
        if (isTable(block)) {
            return modernDefaults ? withoutCellFloors(block) : block;
        }

        // `w:snapToGrid w:val="0"` — the paragraph asked to come off the grid,
        // and comes off it whole: measured, it steps the font's own
        // 11.50 beside a neighbour stepping the grid's 18.00.
        //
        // Asked only of PARAGRAPHS, and after the table above, so the question
        // is only ever put about a block the reader recorded. That leaves the
        // predicate's own default for an unrecorded one unobservable — a
        // mutation flipping it changes nothing and cannot be killed, which is
        // the honest reason it is written where the reader can see the map
        // rather than defended here.
        if (!snaps(block)) {
            return block;
        }

        const style = { ...block.style, gridPitchPx: pitchPx };
        const declaredPx = block.style.lineHeightPx;

        if (undefined === declaredPx) {
            return { ...block, style: { ...style, lineHeightPx: pitchPx, lineRule: 'grid' as LineRule } };
        }

        if ('exact' === block.style.lineRule) {
            return { ...block, style };
        }

        // `auto` resolved to a multiple of the FONT's line and is re-taken
        // against the pitch. The multiple survives that first resolution
        // exactly — it is what turned the font's line into this number — so
        // dividing recovers it rather than guessing at it.
        //
        //  `atLeast` was thought to need nothing here — "its floor is the
        // pitch, a gridded line already takes whole pitches, and a mutation
        // taking max(stated, pitch) as well survived, which is what dead code
        // looks like". MEASURED, and it is not dead code, it is a blind
        // fixture: nothing had ever stated a floor ABOVE the pitch.
        //
        // Printed in an 18.00 grid over 9pt text, baseline to baseline:
        //
        //   body, atLeast 12.00   18.00   the pitch wins        we agree
        //   body, atLeast 24.00   24.00   the floor wins        we give 21.45
        //   cell, atLeast 12.00   18.00   the pitch, in a CELL  we give 12.00
        //   cell, atLeast 24.00   24.00   the floor             we agree
        //
        // Two faults, and the second is the more interesting: a cell is off
        // the grid for its own line rule, and this says it is ON it for the
        // pitch. Not built here — it is a rule about where the grid reaches,
        // which wants its own slice rather than a patch inside this one.
        //
        // And the body fault is NOT in this branch — it is not in this FILE.
        // Taking `max(declared, pitch)` here leaves the 21.45 exactly where it
        // was, and the reader hands the layout precisely what it should: this
        // paragraph arrives as `rule=grid, height=24.00`, and a 21.45 step
        // comes out the other end. The next slice looks at the LAYOUT's grid
        // rule, not at this rescaling.
        // An `atLeast` paragraph stays `atLeast`, with the PITCH as its floor.
        //
        // Calling it `grid` was the fault: the grid rule spends
        // `(height / pitch - 1)` of the grid's own leading at the foot of the
        // paragraph — measured for a MULTIPLE, where 1.5-line spacing means
        // something by 1.5 — and a floor of 24.00 on an 18.00 grid is not a
        // multiple of anything. Read as one it gave back 2.55, which is where
        // the 21.45 came from against the print's 24.00.
        //
        // As `atLeast` both cases fall out: the layout takes the larger of the
        // floor and the line's own text, so a floor under the pitch prints the
        // pitch — 18.00 for a stated 12.00 — and one above it prints itself.
        if ('atLeast' === block.style.lineRule) {
            return {
                ...block,
                style: {
                    ...style,
                    // The floor itself: the layout's grid rule already
                    // takes whole pitches — `max(declared, ceil(natural /
                    // pitch) * pitch)` — so raising it to the pitch here
                    // changes nothing a mutation can find, and this one
                    // really is dead code rather than a blind fixture.
                    lineHeightPx: declaredPx,
                    lineRule: 'grid' as LineRule,
                    gridFloor: true,
                },
            };
        }

        const naturalPx = naturalLineHeightOf(block.runs);
        const rescaledPx = naturalPx > 0 ? (declaredPx / naturalPx) * pitchPx : pitchPx;

        return {
            ...block,
            style: { ...style, lineHeightPx: rescaledPx, lineRule: 'grid' as LineRule },
        };
    });
}

function readPageBorders(sectPr: XmlElement | null): PageBorders | null {
    const element = sectPr?.element('w:pgBorders') ?? null;
    const borders = readBorders(element);

    if (null === element || undefined === borders) {
        return null;
    }

    // `w:offsetFrom` defaults to `text`, which is what a border with no say in
    // the matter gets — and the two are 22.6pt apart on A4 at a one-centimetre
    // margin, so the default is not a detail.
    return {
        borders,
        offsetFrom: 'page' === element.attribute('w:offsetFrom') ? 'page' : 'text',
    };
}

function readFirstPageNumber(sectPr: XmlElement | null): number | null {
    const start = attributeNumber(sectPr?.element('w:pgNumType') ?? null, 'w:start');

    return null !== start && start >= 0 ? start : null;
}

/**
 * `w:pgNumType/@w:fmt` — the numerals this section's page numbers are written
 * in.
 *
 * Folded onto the four this engine can write. Anything else — `ordinal`,
 * `cardinalText`, the East Asian systems — is left as decimal rather than
 * guessed at, which is what the reader does with an unknown field switch too.
 */
function readPageNumberFormat(sectPr: XmlElement | null): NumeralStyle | null {
    const format = sectPr?.element('w:pgNumType')?.attribute('w:fmt') ?? null;

    switch (format) {
        case 'lowerRoman':
            return 'lowerRoman';
        case 'upperRoman':
            return 'upperRoman';
        case 'lowerLetter':
            return 'lowerLetter';
        case 'upperLetter':
            return 'upperLetter';
        case 'decimal':
            return 'decimal';
        default:
            return null;
    }
}

function readGeometry(sectPr: XmlElement | null): PageGeometry {
    const size = sectPr?.element('w:pgSz') ?? null;
    const margin = sectPr?.element('w:pgMar') ?? null;

    const widthTwips = attributeNumber(size, 'w:w') ?? DEFAULT_GEOMETRY.widthTwips;
    // `w:gutter` — the BINDING margin: room left on the edge the pages are
    // stitched or punched at, so the text is not swallowed by the fold. Word
    // writes the attribute on every document and almost always as zero, which
    // is why ignoring it went unnoticed for the whole of this arc.
    //
    // It is part of the LEFT MARGIN rather than a thing of its own, and the
    // measurement says so twice over: a section with 720 twips of gutter began
    // its text at 108.10 against the control's 72.10 — half an inch further in
    // — and one with 1440 began at 144.10, the shift following the value. The
    // writing width comes off with it rather than the text sliding over the
    // right margin: the same paragraph fitted 22 words a line with no gutter,
    // 20 with half an inch and 18 with a whole one.
    //
    // Adding it here, before the columns and the margin are given back, is
    // what carries it to everything measured from the margin — and the header
    // and footer moved with the body in the same print, all four sections
    // drawing their furniture at exactly the body's own left edge.
    const gutterTwips = attributeNumber(margin, 'w:gutter') ?? 0;
    const leftTwips = (attributeNumber(margin, 'w:left') ?? DEFAULT_GEOMETRY.marginTwips)
        + gutterTwips;
    const rightTwips = attributeNumber(margin, 'w:right') ?? DEFAULT_GEOMETRY.marginTwips;
    // Column edges are read from the margin, so the margin is added back here
    // rather than by everything that later asks where a column starts.
    const columns = readColumns(sectPr, widthTwips - leftTwips - rightTwips)
        ?.map((col) => ({ ...col, leftPx: col.leftPx + twipsToPx(leftTwips) }));

    return {
        ...(undefined === columns || null === columns ? {} : { columns }),
        widthPx: twipsToPx(widthTwips),
        heightPx: twipsToPx(attributeNumber(size, 'w:h') ?? DEFAULT_GEOMETRY.heightTwips),
        marginTopPx: twipsToPx(attributeNumber(margin, 'w:top') ?? DEFAULT_GEOMETRY.marginTwips),
        marginRightPx: twipsToPx(rightTwips),
        marginBottomPx: twipsToPx(attributeNumber(margin, 'w:bottom') ?? DEFAULT_GEOMETRY.marginTwips),
        marginLeftPx: twipsToPx(leftTwips),
        headerDistancePx: twipsToPx(attributeNumber(margin, 'w:header') ?? 0),
        footerDistancePx: twipsToPx(attributeNumber(margin, 'w:footer') ?? 0),
    };
}

/**
 * A cell margin, which is written as a child element with its own `w:w`
 * rather than as an attribute of the parent.
 */
/** Include a key only when there is a value — see style-sheet's own `pick`. */
function pick<K extends string, V>(key: K, value: V | null): { [P in K]?: V } {
    return (null === value ? {} : { [key]: value }) as { [P in K]?: V };
}

function marginTwips(margins: XmlElement | null, side: string): number | null {
    return attributeNumber(margins?.element(side) ?? null, 'w:w');
}

/** The `w:val` toggle convention, for the row properties read above. */
function toggleOf(parent: XmlElement | null | undefined, name: string): boolean | null {
    const element = parent?.element(name) ?? null;
    if (null === element) {
        return null;
    }
    const value = element.attribute('w:val');

    return null === value ? true : !('0' === value || 'false' === value || 'off' === value);
}

function attributeNumber(element: XmlElement | null, attribute: string): number | null {
    const raw = element?.attribute(attribute) ?? null;
    if (null === raw) {
        return null;
    }
    const value = Number(raw);

    return Number.isFinite(value) ? value : null;
}

/**
 * Direct formatting on the paragraph itself, which beats anything a style says.
 *
 * Re-read here rather than reusing the style sheet's parser output because the
 * two are merged in a specific order and mixing them up silently inverts the
 * precedence.
 */
function readDirectParagraphProperties(pPr: XmlElement | null | undefined): ParagraphProperties {
    if (null === pPr || undefined === pPr) {
        return {};
    }

    // `w:pBdr` is read HERE rather than by the style sheet: a border is parsed
    // by the same reader a table's is, and that reader lives beside this one.
    const borders = readBorders(pPr.element('w:pBdr'));

    return {
        ...readParagraphProperties(pPr),
        ...(undefined === borders ? {} : { borders }),
    };
}

function readDirectRunProperties(rPr: XmlElement | null | undefined): RunProperties {
    return null === rPr || undefined === rPr ? {} : readRunProperties(rPr);
}
