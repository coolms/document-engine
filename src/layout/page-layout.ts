import { alignLine, type Alignment } from './alignment.js';
import { baselineOffsetPx, type LineRule } from './baseline.js';
import type { NumeralStyle } from '../text/numerals.js';
import { hasField, resolveFields, type FieldContext } from './fields.js';
import {
    isFloatingBox,
    lineBoxAt,
    type Float,
    type LineBox,
    type PlacedFloat,
} from './float.js';
import type { InlineImage, InlineShape } from './image.js';
import type { TrueTypeFont } from '../font/truetype-font.js';
import {
    advanceOf,
    borderRoomOf,
    breakIntoLines,
    runsFrom,
    type Line,
    type RubyGloss,
    type StyledRun,
    type TabStop,
} from './line-breaker.js';
import {
    headerRows,
    isTable,
    measureRow,
    fitTable,
    ruleAbove,
    ruleBelow,
    measureTable,
    placeRow,
    resolveTableGeometry,
    splitRow,
    stackBlocks as stackBlocksIn,
    translateLine,
    translateRow,
    verticalSpans,
    withSide,
    type Block,
    type LineMetrics,
    type MeasuredRow,
    type PlacedRow,
    type StackedBlocks,
    type Table,
} from './table-layout.js';

export { alignLine } from './alignment.js';
import {
    borderRoom as roomForSide,
    borderStandoff as standoffOf,
    sameBorders,
    strongerBorder,
    type BorderSide,
    type BoxBorders,
    type PlacedParagraphBorder,
} from './borders.js';

export { resolveCellBorders, strongerBorder } from './borders.js';
export type { BorderSide, BorderStyle, BoxBorders, CellPosition } from './borders.js';
export type { Alignment } from './alignment.js';
export { baselineOffsetPx, EXACT_BASELINE_RATIO } from './baseline.js';
export { isFloatingBox, lineBoxAt } from './float.js';
export type {
    BoxInset,
    Float,
    FloatingBox,
    FloatingImage,
    FloatPosition,
    LineBox,
    PlacedFloat,
    RelativeTo,
    WrapMode,
} from './float.js';
export type { BaselineMetrics, LineRule } from './baseline.js';

export { fitTable, isTable, verticalSpans } from './table-layout.js';
export type { PlacedParagraphBorder } from './borders.js';
export type {
    Block,
    CellMargins,
    CellTextDirection,
    CellVerticalAlign,
    TableAlignment,
    PlacedCell,
    PlacedRow,
    RowHeightRule,
    Table,
    TableCell,
    TableRow,
    VerticalMerge,
} from './table-layout.js';

/** What a run of `keepNext` paragraphs needs in order to stay together. */
interface KeepGroup {
    /** The last paragraph the keep binds, so its members are not asked twice. */
    readonly end: number;
    /**
     * The room the group needs on one page, or null where the keep cannot apply
     * and must not cost a page break: a lone paragraph at the end of the
     * document, or one in front of a block that opens a page of its own anyway.
     */
    readonly heightPx: number | null;
}

/** A paragraph's lines, measured where they are going to sit. */
interface FlowedParagraph {
    readonly lines: Line[];
    /** The nominal height a float was placed against, before the breaking. */
    readonly lineHeight: number;
    /** The box each line actually got, one per line. */
    readonly metrics: readonly LineMetrics[];
    /** The full column, before any float narrows an individual line. */
    readonly column: LineBox;
    readonly boxOf: (lineIndex: number, topPx: number) => LineBox;
}

/**
 * Flowing blocks onto pages, breaking to a new page when they no longer fit.
 *
 * This is the piece that makes a page break mean something: content that
 * overruns a page continues on the next one instead of the page growing.
 */

export interface PageGeometry {
    readonly widthPx: number;
    readonly heightPx: number;
    readonly marginTopPx: number;
    readonly marginRightPx: number;
    readonly marginBottomPx: number;
    readonly marginLeftPx: number;
    /** `w:pgMar/@w:header` — page edge to the TOP of the header. */
    readonly headerDistancePx?: number;
    /** `w:pgMar/@w:footer` — page edge to the BOTTOM of the footer. */
    readonly footerDistancePx?: number;
    /**
     * `w:cols` — the columns text flows down before it reaches the next page.
     *
     * Absolute rather than a count and a gap, because Word can state each
     * column's own width (`w:equalWidth="0"`), and a count could not say that.
     * Absent or a single entry is an ordinary page.
     */
    readonly columns?: readonly PageColumn[];
}

/** One column of a `w:cols` section. */
export interface PageColumn {
    /** Left edge on the page, the margin already taken in. */
    readonly leftPx: number;
    readonly widthPx: number;
}

export interface ParagraphStyle {
    readonly spaceBeforePx: number;
    readonly spaceAfterPx: number;
    /** Absent is `left`, which is what a paragraph declaring no `w:jc` gets. */
    readonly alignment?: Alignment;
    /** Exact line height. When absent, the font's natural single spacing wins. */
    readonly lineHeightPx?: number;
    /**
     * How that height was ARRIVED at — `w:spacing/@w:lineRule`.
     *
     * Kept even though the flow only needs the number, because two lines of
     * identical height put their baseline in different places depending on
     * which rule produced it. Absent means `auto`, as it does in Word.
     */
    readonly lineRule?: LineRule;
    /**
     * The `w:docGrid` pitch this paragraph's section is ruled to, where it has
     * one.
     *
     * Beside {@link lineHeightPx} rather than folded into it because the two
     * answer different questions once a paragraph states spacing of its own:
     * the height is the room the line ASKS for and the pitch is what its
     * leading is measured against. Measured: under an 18pt grid, a
     * paragraph at 1.5 lines steps 27.00 — one and a half PITCHES, not one and
     * a half of the font's 11.50 — while its baseline stays 12.62 down, the
     * natural line centred in a single pitch with the rest of the room below.
     */
    readonly gridPitchPx?: number;
    /**
     * The gridded height is a FLOOR — `atLeast` — rather than a multiple.
     *
     * A grid paragraph gives leading back at its foot in proportion to the
     * multiple it asked for, which is meaningless for a floor: 24.00 on an
     * 18.00 grid is not 1.33 of anything, and reading it as one printed a
     * 21.45 step where the page steps 24.00. The RULE stays
     * `grid`, since the baseline still sits in the middle of its grid
     * line; only the giving back is skipped.
     */
    readonly gridFloor?: boolean;
    /**
     * `w:suppressLineNumbers` — this paragraph's lines are not counted.
     *
     * On the style rather than worked out per line because the numbering runs
     * over PLACED pages, long after the paragraph that asked for it went past.
     */
    readonly suppressLineNumbers?: boolean;
    readonly indentLeftPx: number;
    readonly indentRightPx: number;
    /** Extra indent on the FIRST line only; negative gives a hanging indent. */
    readonly indentFirstLinePx: number;
    /** Start this paragraph on a new page whatever room is left. */
    readonly pageBreakBefore: boolean;
    /**
     * Keep every line together on one page. Word calls it "Keep lines
     * together"; it is off by default and turned on for headings and captions.
     */
    readonly keepLinesTogether: boolean;
    /**
     * `w:keepNext` — keep this paragraph on the page that holds the block after
     * it. Word calls it "Keep with next" and turns it on for headings, so a
     * heading is never stranded at the foot of a page above its own text.
     *
     * Consecutive flagged paragraphs CHAIN: a heading kept with a subheading
     * kept with a body paragraph all travel together.
     */
    readonly keepWithNext: boolean;
    /**
     * Widow and orphan control for THIS paragraph, overriding the document-wide
     * setting. Word stores `w:widowControl` per paragraph, so a document can
     * switch it off for one block and leave it on everywhere else; a
     * document-wide flag alone cannot express that.
     *
     * Absent means "follow the document".
     */
    readonly widowControl?: boolean;
    /**
     * `w:pBdr` — the paragraph's OWN border: a rule under a heading, a box
     * round a call-out.
     *
     * Consecutive paragraphs carrying the same border are one box, which is
     * why this is compared between them rather than drawn per paragraph.
     */
    readonly borders?: BoxBorders;
    /** Explicit tab stops, from the paragraph's left indent, ascending. */
    readonly tabStops?: readonly TabStop[];
    /** `w:decimalSymbol` — what a decimal stop lines its numbers up on. */
    readonly decimalSymbol?: string;
    /**
     * The repeating stop past the last explicit one.
     *
     * Word's default is half an inch, which is 48px here — and a tab that
     * advanced by its GLYPH instead would move 8px in Carlito and 12 in
     * Liberation Serif, neither of which is a column.
     */
    readonly defaultTabPx?: number;
}

/** Word's default tab stop: half an inch, in CSS pixels. */
export const DEFAULT_TAB_PX = 48;

export const DEFAULT_PARAGRAPH_STYLE: ParagraphStyle = {
    spaceBeforePx: 0,
    spaceAfterPx: 0,
    indentLeftPx: 0,
    indentRightPx: 0,
    indentFirstLinePx: 0,
    pageBreakBefore: false,
    keepLinesTogether: false,
    keepWithNext: false,
};

export interface ListMarker {
    readonly run: StyledRun;
    /**
     * How far LEFT of the paragraph's left indent the marker is drawn — the
     * hanging indent.
     *
     * Its own field because the marker sits there and the TEXT does not: a
     * list's first line starts at the left indent like every other line, and
     * giving it the negative first-line indent instead draws the text on top
     * of the bullet.
     */
    readonly offsetPx: number;
}

export interface Paragraph {
    readonly runs: readonly StyledRun[];
    readonly style: ParagraphStyle;
    /**
     * A list marker — "1.", "a)", a bullet glyph — drawn in the hanging
     * indent beside the FIRST line.
     *
     * Not part of the runs, because it is not part of the text: it is never
     * measured for wrapping, never selected, and never breaks a line.
     */
    readonly marker?: ListMarker;
    /**
     * Pictures anchored to this paragraph that text flows AROUND.
     *
     * On the paragraph rather than in its runs because that is what they are
     * anchored to: a float is placed relative to the paragraph's own top, and
     * the lines of this paragraph and the ones after it step aside for it.
     */
    readonly floats?: readonly Float[];
}

export interface PlacedLine {
    readonly line: Line;
    /** Left edge, margins, indents and alignment already applied. */
    readonly xPx: number;
    /**
     * Extra width to add to EACH space on this line — CSS `word-spacing`.
     *
     * Zero unless the paragraph is justified. Carried as one number rather
     * than as re-measured pieces because that is all a renderer needs, and
     * because the slack has to be divided once: dividing it again per piece
     * would let rounding accumulate across a line.
     */
    readonly wordSpacingPx: number;
    /** Top edge, relative to the page. */
    readonly yPx: number;
    readonly heightPx: number;
    /**
     * How far below {@link yPx} the baseline sits — where the glyphs rest.
     *
     * On the line rather than left to the renderer, because working it out
     * needs the line RULE and the font's descent, and a renderer given only
     * a box would have to re-derive both.
     */
    readonly baselinePx: number;
    /** Index of the paragraph this line came from, for hit-testing and editing. */
    readonly paragraphIndex: number;
    /** Index of the line WITHIN its paragraph. */
    readonly lineIndex: number;
    /**
     * The list marker drawn beside this line, on the FIRST line of a list item
     * and nowhere else.
     *
     * Placed here rather than left to a renderer to fetch from the block,
     * because a page is meant to be drawable from itself: a renderer that had
     * to reach back into the document to find the bullet would need the blocks,
     * the paragraph index and the hanging-indent rule as well.
     *
     * Its x is NOT relative to this line. A marker sits in the hanging space
     * left of the paragraph's own indent, so a centred or indented first line
     * moves while the bullet stays put.
     */
    readonly marker?: PlacedMarker;
}

export interface PlacedMarker {
    readonly run: StyledRun;
    /** Left edge of the marker, in the same space as {@link PlacedLine.xPx}. */
    readonly xPx: number;
}


/**
 * A header or footer, placed on one page.
 *
 * Separate from the page's own lines because it is not part of the flow: it
 * repeats, it is not selectable as body text, and its coordinates come from the
 * paper rather than from the cursor. Folding it into `lines` would make every
 * consumer re-separate it, and the paragraph indices would collide with the
 * body's.
 */
/**
 * The rule Word draws above a page's footnotes: two inches, at the margin.
 *
 * MEASURED at last, where this was an uncited claim that happened to
 * be right: the print runs it from 28.350 to 172.350 on a page with 1cm
 * margins — 144.00pt, two inches exactly, starting ON the margin. `audit:claims`
 * never asked because 192 is a whole number, which is worth remembering about
 * that guard: it catches a fraction nobody measured, not a round number nobody
 * measured.
 */
export const FOOTNOTE_RULE_WIDTH_PX = 192;

/** What the rule's own line is worth when the note holds no paragraph to ask. */
const DEFAULT_SEPARATOR_HEIGHT_PX = 15.333333333333334;

/**
 * How far above the notes that rule sits.
 *
 * Measured against LibreOffice at 4.1pt, which is 5.47px, and the same whether
 * the page carries one note or two.
 */
export const FOOTNOTE_RULE_GAP_PX = 4.1 * 96 / 72;

/** A page's footnotes, already laid out at the foot of it. */
export interface PlacedFootnotes {
    /** Top of the reserved area, which is where the rule is drawn from. */
    readonly separatorYPx: number;
    readonly separatorLeftPx: number;
    readonly separatorWidthPx: number;
    /** Every note's lines, in PAGE coordinates and in reference order. */
    readonly lines: readonly PlacedLine[];
}

export interface PlacedFurniture {
    /** Its lines, in PAGE coordinates. */
    readonly lines: readonly PlacedLine[];
    /** Rows of any table it holds, in PAGE coordinates. */
    readonly rows: readonly PlacedRow[];
    /**
     * Boxes round its bordered paragraphs, in PAGE coordinates.
     *
     * A header is stacked by the same code a cell is, so it gathers these the
     * same way — and dropping them here is what left a bordered heading in a
     * header drawing nothing. LibreOffice draws it: 14.45pt round an 11.5pt
     * line for a three-point rule, standing 1.5 outside the text either side.
     */
    readonly paragraphBorders: readonly PlacedParagraphBorder[];
    /** Top edge on the page. */
    readonly topPx: number;
    readonly heightPx: number;
}

export interface Page {
    /**
     * Text belonging to the page itself. Lines inside a table cell are NOT
     * here — they live on their cell, where their position and their column
     * are known together.
     */
    readonly lines: readonly PlacedLine[];
    /** Table rows on this page, including headers repeated onto it. */
    readonly rows: readonly PlacedRow[];
    /**
     * This page's header, already positioned — absent when it has none.
     *
     * A header GROWS DOWNWARDS from its distance off the top edge, and a footer
     * grows UPWARDS from its distance off the bottom one. Placing the footer
     * downwards instead puts a two-line footer off the bottom of the paper,
     * which is why the two are computed differently rather than shared.
     */
    readonly header?: PlacedFurniture;
    readonly footer?: PlacedFurniture;
    /** Set when a reference on this page put a note at the foot of it. */
    readonly footnotes?: PlacedFootnotes;
    /** The boxes round paragraphs that asked for one. */
    readonly paragraphBorders: readonly PlacedParagraphBorder[];
    /** `w:pgBorders`, already resolved against this page's own paper. */
    readonly pageBorder?: PlacedParagraphBorder;
    /** `w:lnNumType`, counted and placed. Empty where the section asks for none. */
    readonly lineNumbers: readonly PlacedLineNumber[];
    /**
     * Floating pictures on this page, already placed.
     *
     * On the PAGE rather than on a line, because a float belongs to neither:
     * it is anchored to a paragraph but positioned against the page, and the
     * lines it displaces may come from several paragraphs after its anchor.
     */
    readonly floats: readonly PlacedFloat[];
    /**
     * Where this page falls in the printed document, counting from one.
     *
     * On the page rather than left to the caller's index because a section may
     * begin part-way through: a header that alternates by parity needs the
     * DOCUMENT's number, and a page that only knows its position within its own
     * section cannot supply it.
     */
    readonly pageNumber: number;
    /**
     * The paper THIS page is on.
     *
     * Per page rather than per document because a section break may change
     * it: a landscape table page inside a portrait report is one document
     * with two page sizes, and a renderer given only the first would draw
     * the second at the wrong size.
     */
    readonly geometry: PageGeometry;
    /** Which section this page belongs to. */
    readonly sectionIndex: number;
}

/**
 * `w:pgBorders` — a box round the whole page, and what it is measured from.
 *
 * `page` counts from the edge of the PAPER inward and `text` from the writing
 * area outward, which put the same border 22.6pt apart on A4 with a
 * one-centimetre margin. Either way the `w:space` is the CLEAR GAP between the
 * rule and the thing it is measured from — measured against LibreOffice, a
 * three-point rule 24 points from the paper drew its centre at 25.5.
 */
/**
 * `w:lnNumType` — numbers printed down the margin beside the body's lines.
 *
 * Measured against LibreOffice: the number's RIGHT edge lands `distance` in
 * from the writing area, an empty paragraph is a line and gets one, and a line
 * inside a TABLE does not.
 */
export interface LineNumbering {
    /** `w:countBy` — a number is printed where the count divides by this. */
    readonly countBy: number;
    /**
     * `w:start`, which LibreOffice adds to the count rather than starting at.
     *
     * Stated as five with `countBy` two, it printed 6, 8 and 10 against lines
     * one, three and five — so the first line is `start + 1`, not `start`.
     */
    readonly start: number;
    /** `w:distance`; absent, LibreOffice left half a centimetre. */
    readonly distancePx: number;
    readonly restart: 'newPage' | 'newSection' | 'continuous';
}

/** One number, in the font it is drawn in, at the place it is drawn. */
export interface PlacedLineNumber {
    readonly run: StyledRun;
    /** The LEFT edge: the right-alignment is already resolved. */
    readonly xPx: number;
    readonly baselinePx: number;
}

export interface PageBorders {
    readonly borders: BoxBorders;
    readonly offsetFrom: 'page' | 'text';
}

export interface Section {
    readonly blocks: readonly Block[];
    readonly geometry: PageGeometry;
    /**
     * `w:pgNumType/@w:start` — the number this section's first page PRINTS.
     *
     * Front matter numbered i, ii, iii and a body that begins again at 1 is
     * two sections and one document: without this the body's first page is
     * numbered by how many pages came before it, which is the one thing the
     * document said it is not.
     */
    readonly firstPageNumber?: number;
    /**
     * `w:pgNumType/@w:fmt` — how this section writes its page numbers.
     *
     * Measured: a section saying `lowerRoman` printed i and ii, and
     * the section after it printed 1. It is the ordinary way a document asks
     * for roman front matter — far commoner than a `PAGE \* roman` switch on
     * each field, which is all this engine read before.
     */
    readonly pageNumberFormat?: NumeralStyle;
    /** `w:pgBorders` — drawn on every page of the section. */
    readonly pageBorders?: PageBorders;
    /** `w:lnNumType` — numbers down the margin of every page of the section. */
    readonly lineNumbering?: LineNumbering;
    /**
     * The document's footnotes, by the id its references name.
     *
     * On the SECTION rather than left to the caller because a caller that
     * forgot to hand them over would lose them silently — which is the whole of
     * what was wrong before they were read at all.
     */
    readonly footnotes?: ReadonlyMap<number, readonly Block[]>;
    /** The writing area of page N WITHIN this section. */
    readonly contentBox?: (pageIndex: number, pageNumber: number) => ContentBox;
    /**
     * The blocks of the header this page draws, or nothing when it draws none.
     *
     * Blocks rather than a named variant, so this layer never learns what a
     * variant is: choosing between a first-page, even-page and default header
     * is a WordprocessingML rule, and the layout only has to put what it is
     * given in the right place.
     */
    readonly headerFor?: FurnitureSource;
    readonly footerFor?: FurnitureSource;
    /**
     * Which page this section has to begin on.
     *
     * `evenPage` and `oddPage` are how a chapter always starts on a
     * right-hand page: if the parity is wrong a BLANK page is inserted to
     * reach it, and that blank page is real — it is printed, and it counts.
     */
    readonly startsOn?: 'nextPage' | 'evenPage' | 'oddPage';
}

/**
 * Which blocks a page's header or footer is made of.
 *
 * Given BOTH numbers because Word needs both and they diverge. "Different first
 * page" means the first page of the SECTION, while "different odd and even"
 * counts PRINTED pages — so a section opening on document page four is its own
 * first page and an even page at once, and one index cannot say which.
 *
 * Verified against LibreOffice: with both settings on, such a page draws the
 * new section's FIRST header; with `w:titlePg` absent it draws the new
 * section's EVEN header, not its default one.
 */
export type FurnitureSource = (
    pageIndexInSection: number,
    pageNumber: number,
) => readonly Block[] | undefined;

export interface ContentBox {
    /** Top of the writing area, below anything the header occupies. */
    readonly topPx: number;
    /** Bottom of it, above anything the footer occupies. */
    readonly bottomPx: number;
}

export interface LayoutOptions {
    /** `w:pgNumType/@w:fmt` — how the section being laid writes its numbers. */
    readonly pageNumberFormat?: NumeralStyle;
    /**
     * Each footnote's blocks, by the id its references name.
     *
     * Handed in rather than reached for, because a note belongs to the page its
     * reference lands on and only the flow knows which page that is.
     */
    readonly footnotes?: ReadonlyMap<number, readonly Block[]>;
    /**
     * Word's "Widow/Orphan control", which is **ON by default** — so this
     * defaults to true as well. Leaving it off would make our page breaks
     * disagree with the .docx for any paragraph that straddles a boundary,
     * which is most of them in a long document.
     *
     * ⚠️ THE REFERENCE DISAGREES, and only here. LibreOffice honours
     * `w:widowControl` when a document states it — asked for outright, both a
     * stranded first line and a stranded last one moved the whole paragraph on
     * — but with the element ABSENT it allows both: a printed page ended with
     * one line of a three-line paragraph, and another left the third line
     * alone overleaf.
     *
     * Kept as true because the thing being modelled is WORD's layout, and
     * LibreOffice is the reference this engine can measure rather than the
     * authority on what a `.docx` means. Two witnesses agree against it: the
     * format states the default is true, and Word applies it — it is the
     * "Widow/Orphan control" tick. The divergence is confined to SILENCE: a
     * document that states the element either way is laid out identically by
     * both. Not one document in this corpus states it, including the
     * Word-authored one, so it decides the pagination of all of them.
     */
    readonly widowOrphanControl?: boolean;
    /**
     * The writing area of page N, when it is not simply the margins.
     *
     * Asked PER PAGE because it varies: a document with a different first-page
     * header has a different amount of room on page one, and a header taller
     * than its margin pushes the text down on the pages that have it. One box
     * for the whole document would be right for neither.
     */
    readonly contentBox?: (pageIndex: number, pageNumber: number) => ContentBox;
    /** See {@link Section.headerFor}. */
    readonly headerFor?: FurnitureSource;
    readonly footerFor?: FurnitureSource;
    /**
     * The printed number of this run's FIRST page. One unless other sections
     * came before it — `layoutSections` supplies the running count.
     */
    readonly firstPageNumber?: number;
    /**
     * What a `NUMPAGES` field resolves to.
     *
     * Supplied rather than counted, because this run of pages need not be the
     * whole document: a section laid out on its own knows how many pages IT
     * has, and the footer wants the document's total. Absent falls back to this
     * run's own count, which is right for a single-section document.
     */
    readonly totalPages?: number;
}

/**
 * Lay paragraphs out onto pages.
 *
 * ## Spacing is additive, as Word does it
 *
 * The space after one paragraph and before the next ADD, they do not collapse
 * the way CSS margins do. Collapsing them would shorten every gap in the
 * document and eventually move a page break.
 *
 * ## Widow and orphan control
 *
 * A paragraph is never split so as to leave one line alone. An ORPHAN is its
 * first line stranded at the foot of a page — the whole paragraph moves on. A
 * WIDOW is its last line stranded at the head of the next — one more line is
 * pushed over to join it. Both are on by default because they are on by default
 * in Word, and this engine's purpose is to agree with the file it produces.
 *
 * A paragraph too tall for a page is exempt: nothing can be done for it, and
 * refusing to split it would loop forever.
 */
export function layoutPages(
    blocks: readonly Block[],
    geometry: PageGeometry,
    options: LayoutOptions = {},
): Page[] {
    const widowOrphan = options.widowOrphanControl ?? true;
    const footnotesById = options.footnotes ?? new Map<number, readonly Block[]>();

    // The whole writing width, which is what furniture spans however many
    // columns the body is flowing down.
    const writingWidth = geometry.widthPx - geometry.marginLeftPx - geometry.marginRightPx;
    const columns: readonly PageColumn[] = undefined === geometry.columns
        || 0 === geometry.columns.length
        ? [{ leftPx: geometry.marginLeftPx, widthPx: writingWidth }]
        : geometry.columns;
    const boxOf = options.contentBox ?? ((): ContentBox => ({
        topPx: geometry.marginTopPx,
        bottomPx: geometry.heightPx - geometry.marginBottomPx,
    }));
    const firstPageNumber = options.firstPageNumber ?? 1;
    const numberOf = (pageIndex: number): number => firstPageNumber + pageIndex;

    interface Sheet {
        lines: PlacedLine[];
        rows: PlacedRow[];
        floats: PlacedFloat[];
        /** Notes referenced from this page, in the order they were referenced. */
        notes: number[];
        borders: PlacedParagraphBorder[];
        /**
         * How much of the PREVIOUS page's notes had to be carried here.
         *
         * A note taller than the room its page can spare is continued rather
         * than drawn over the text — so the page it continues onto owes that
         * much before it owes anything of its own.
         */
        carriedPx: number;
    }

    const pages: Sheet[] = [
        { lines: [], rows: [], floats: [], notes: [], carriedPx: 0, borders: [] },
    ];

    const current = (): Sheet => pages[pages.length - 1]!;
    // Emptiness has to count ROWS as well as lines, or a table at the top of
    // a page reads as 'nothing here yet' and the next break opens a blank one.
    // ...and where the COLUMN being filled began, so a break decision asks
    // about the column rather than about the page it happens to sit on.
    let columnStart = { lines: 0, rows: 0 };
    const isEmpty = (): boolean => current().lines.length === columnStart.lines
        && current().rows.length === columnStart.rows;
    const box = (): ContentBox => boxOf(pages.length - 1, numberOf(pages.length - 1));
    // What is left of the page once its notes have had their room. Everything
    // that asks where the page ENDS asks this, so a footnote pushes text off
    // the page by the same path a short page does.
    const bottom = (): number => box().bottomPx - reservedFor(current());
    /** The page's own foot, which the notes themselves sit against. */
    const pageFoot = (): number => box().bottomPx;
    // The room on THIS page, which the pages either side of it need not share.
    const contentHeight = (): number => box().bottomPx - box().topPx;

    /**
     * A note's own height, measured once and kept.
     *
     * The same note can be asked about many times over while the page decides
     * whether the line referencing it fits, and a note is a column of blocks
     * like any other — measuring it each time would flow the whole thing again
     * for every candidate line.
     */
    const noteHeights = new Map<number, number>();
    const noteHeight = (id: number): number => {
        const known = noteHeights.get(id);
        if (undefined !== known) {
            return known;
        }

        const blocks = footnotesById.get(id) ?? [];
        const measured = stackBlocks(blocks, writingWidth).heightPx;
        noteHeights.set(id, measured);

        return measured;
    };

    /**
     * The room the foot of the page owes its notes.
     *
     * The rule sits on a LINE of its own above them — measured against
     * LibreOffice, a page carrying one note held twelve body lines where an
     * unfootnoted one held fourteen, which is the note AND the line the rule
     * stands on. Nothing at all is owed while the page carries no notes.
     */
    /** Everything a page's notes want: what it carried in, and its own. */
    const notesHeight = (sheet: { notes: readonly number[]; carriedPx: number }): number => {
        if (0 === sheet.notes.length && 0 === sheet.carriedPx) {
            return 0;
        }

        let total = sheet.carriedPx;
        for (const id of sheet.notes) {
            total += noteHeight(id);
        }

        return total + separatorHeightPx(sheet.notes[0]);
    };

    /**
     * What the page can actually GIVE them, which is not always what they want.
     *
     * A note taller than the page cannot have all of it: measured against
     * LibreOffice, the referencing line stayed where it was and the notes took
     * everything below it, the rest continuing overleaf. So the reserve is
     * capped at the room below the text already placed — never at nothing,
     * because a note area of no height would carry the whole note forever.
     */
    const reservedFor = (
        sheet: { notes: readonly number[]; carriedPx: number },
        atY = cursorY,
    ): number => {
        const wanted = notesHeight(sheet);

        // Notes that would fit on a page of their own are worth moving text
        // for: the line that brings them travels to where they both fit, which
        // is what LibreOffice did with a note of one line. Notes that would
        // NOT fit anywhere are given the room below the text instead, and what
        // is left of them is carried overleaf — where a page that kept holding
        // out for room it can never have would move the line forever.
        return wanted <= contentHeight() ? wanted : Math.max(0, pageFoot() - atY);
    };

    /** The line the rule stands on, which is one line of the note's own text. */
    const separatorHeightPx = (firstNote: number | undefined): number => {
        if (undefined === firstNote) {
            return DEFAULT_SEPARATOR_HEIGHT_PX;
        }

        const [first] = footnotesById.get(firstNote) ?? [];

        return undefined === first || isTable(first)
            ? DEFAULT_SEPARATOR_HEIGHT_PX
            : paragraphLineHeight(first);
    };

    /**
     * The box being built, if a bordered paragraph is open.
     *
     * A run of paragraphs carrying the SAME border is one box, so it is
     * extended paragraph by paragraph and closed when the border changes, when
     * an unbordered paragraph follows, or when the page does.
     */
    let openBorder: {
        borders: BoxBorders;
        leftPx: number;
        rightPx: number;
        topPx: number;
        bottomPx: number;
        opensHere: boolean;
        innerYPx: number[];
    } | null = null;

    /**
     * Finish the open box, if there is one.
     *
     * A box always CLOSES where it stops — at the end of its run of
     * paragraphs, and at a page or column break alike. Measured against
     * LibreOffice: a bordered paragraph split over two pages drew FOUR rules on
     * each of them, a complete box either side, not one outline left open at
     * the foot and picked up at the head. The tempting analogy is a table row,
     * which splits without drawing the edge between its halves — a paragraph's
     * border does not behave like one, and reading it that way leaves the two
     * inner edges undrawn.
     */
    const closeBorder = (): void => {
        if (null === openBorder) {
            return;
        }

        current().borders.push({ ...openBorder, closesHere: true });
        openBorder = null;
    };

    const standoff = standoffOf;

    /**
     * The room a side takes outside the text: its space plus its whole width.
     *
     * The rule is centred half a width outside the box, so its OUTER edge —
     * what the text above or below has to clear — is a full width out.
     * Measured off LibreOffice, where a bordered paragraph under an ordinary
     * one stepped 12.5pt instead of 11.5 for a one-point rule, and 18.5 once
     * six points of `w:space` were added.
     */
    const roomFor = (side: BorderSide | undefined): number =>
        undefined === side ? 0 : (side.spacePx ?? 0) + side.widthPx;

    /** Close the open box and give its bottom rule the room it takes. */
    const endBorderRun = (): void => {
        const bottom = openBorder?.borders.bottom;

        closeBorder();
        cursorY += roomFor(bottom);
    };

    /**
     * Grow the open box to take in a line, opening one if it has to.
     *
     * Line by line rather than paragraph by paragraph, because a paragraph can
     * span two pages and the two halves of its box are drawn in different
     * places. The edges are the LINE BOX's own, with each rule centred half a
     * width outside and `w:space` pushing it further out again — measured off
     * LibreOffice, where six points of space moved the box six points out on
     * every side and left the text exactly where it was.
     */
    const extendBorder = (
        borders: BoxBorders,
        style: ParagraphStyle,
        topPx: number,
        bottomPx: number,
    ): void => {
        const edges = {
            leftPx: contentLeft() + style.indentLeftPx - standoff(borders.left),
            rightPx: contentLeft() + contentWidth() - style.indentRightPx
                + standoff(borders.right),
        };

        // Whether the run goes on is settled at the paragraph boundary, above,
        // so by the time a line is placed a box is either open and this line's
        // own or not open at all. There is nothing left to compare here.
        if (null !== openBorder) {
            openBorder = { ...openBorder, bottomPx: bottomPx + standoff(borders.bottom) };

            return;
        }

        openBorder = {
            borders,
            ...edges,
            topPx: topPx - standoff(borders.top),
            bottomPx: bottomPx + standoff(borders.bottom),
            opensHere: true,
            innerYPx: [],
        };
    };

    let cursorY = box().topPx;
    // The last block whose keep group has been given its page.
    let keepDecidedThrough = -1;

    let columnIndex = 0;
    const column = (): PageColumn => columns[columnIndex]!;
    const contentWidth = (): number => column().widthPx;
    const contentLeft = (): number => column().leftPx;

    const newPage = (): void => {
        // Whatever this page could not give its notes goes with them.
        const owed = notesHeight(current());
        const carriedPx = Math.max(0, owed - reservedFor(current()));

        closeBorder();
        pages.push({ lines: [], rows: [], floats: [], notes: [], carriedPx, borders: [] });
        columnIndex = 0;
        cursorY = box().topPx;
        columnStart = { lines: 0, rows: 0 };
    };

    /**
     * Move on when what is being placed will not fit below the cursor.
     *
     * The next COLUMN, if the section has one left — a page break only comes
     * once the last column is full. That is the whole of what makes a
     * multi-column section: everything else measures against the column it is
     * in, and nothing else needs to know how many there are.
     */
    /** Whether what follows this column is a column of another width. */
    const nextColumnDiffers = (): boolean => {
        const next = columns[columnIndex + 1] ?? columns[0]!;

        return next.widthPx !== column().widthPx;
    };

    const nextColumn = (): void => {
        if (columnIndex + 1 >= columns.length) {
            newPage();

            return;
        }

        closeBorder();
        columnIndex++;
        cursorY = box().topPx;
        columnStart = { lines: current().lines.length, rows: current().rows.length };
    };

    /**
     * Break a paragraph into lines as if its first one began at `startY`.
     *
     * `startY` reaches only the floats, and it has to: which lines a float
     * displaces depends on where those lines SIT, so a paragraph measured at
     * the wrong height wraps around the wrong part of a picture.
     *
     * A line sits at `startY` plus the heights of the lines BEFORE it, which
     * the breaker knows by the time it asks — it used to be `N * lineHeight`,
     * and a paragraph whose first line carried a picture then walked the band
     * three times too fast.
     *
     * The line's own extent is still the paragraph's nominal height, because
     * that is the one number not yet known when the question is asked: a line's
     * height depends on what lands on it, and what lands on it depends on this
     * answer. It only matters for a line that starts ABOVE a float and reaches
     * into it, where the nominal over-states the reach.
     *
     * A paragraph that breaks across a page displaces the lines after the break
     * by their position in the PARAGRAPH rather than on the page — a float is
     * anchored near its paragraph's start, so its band reaches the next page
     * only when it is nearly a page tall.
     */
    const flowParagraph = (paragraph: Paragraph, startY: number): FlowedParagraph => {
        const style = paragraph.style;
        const available = contentWidth() - style.indentLeftPx - style.indentRightPx;
        const column: LineBox = {
            leftPx: contentLeft() + style.indentLeftPx,
            widthPx: available,
        };
        const lineHeight = paragraphLineHeight(paragraph);

        // The LEFT is read when the line is placed, not when the paragraph was
        // measured: a paragraph that fills one column and carries on into the
        // next is drawn where that next column is. The WIDTH is the one it was
        // broken at, so a section whose columns differ in width keeps the lines
        // it already has rather than re-breaking them halfway down.
        const boxOf = (_lineIndex: number, topPx: number): LineBox => lineBoxAt(
            { leftPx: contentLeft() + style.indentLeftPx, widthPx: available },
            startY + topPx,
            lineHeight,
            current().floats,
        );
        const { lines, metrics } = measureParagraph(paragraph, available, boxOf);

        return { lines, lineHeight, metrics, column, boxOf };
    };

    /**
     * How tall a table's first row is.
     *
     * A `keepNext` paragraph in front of a table is kept with that FIRST ROW,
     * not with the whole table: measured against LibreOffice, the rest of the
     * table flowed onto the next page and left the heading with row one.
     */
    const firstRowHeight = (table: Table): number => {
        const first = table.rows[0];
        if (undefined === first) {
            return 0;
        }

        // The spans matter even here: a merged cell is measured against its
        // whole span and so cannot set the height of the row it starts in.
        // Only the LAST row of a span ever grows, and that is never row one.
        return measureRow(
            first,
            table,
            { firstRow: true, lastRow: 1 === table.rows.length },
            (blocks, widthPx) => stackBlocksIn(blocks, widthPx, measureParagraph),
            verticalSpans(table)[0],
        ).heightPx;
    };

    /**
     * The keep group that starts at `from`: what has to share one page.
     *
     * `w:keepNext` binds a paragraph to the block after it, and consecutive
     * flagged paragraphs CHAIN, so the group runs on until a block that carries
     * no keep of its own ends it. It is not symmetric — measured against
     * LibreOffice, every BOUND paragraph moves whole while the block that ends
     * the group need only get its FIRST line onto the page, or its first row
     * when it is a table. A heading kept with a long paragraph therefore needs
     * one line of that paragraph, not all of it.
     *
     * With nothing after it at all the group is kept with its OWN last
     * paragraph, because a heading kept with a subheading still has to sit
     * beside it even where the subheading ends the document.
     *
     * The measurement cannot see floats the group has yet to anchor, since none
     * of it is placed yet — a keep group that also carries a floating picture
     * can still be split.
     */
    const keepGroup = (from: number, atY: number): KeepGroup => {
        // Where the chain ends, and what it is kept WITH, before anything is
        // measured. Which paragraphs have to fit WHOLE depends on the answer:
        // with nothing after it, a group is kept with its own last paragraph,
        // and that one then needs only its first line like any other follower.
        let end = from;
        while (end + 1 < blocks.length) {
            const next = blocks[end + 1]!;
            // A table carries no keep of its own, and a block that opens a page
            // of its own cannot be pulled back onto this one — either ends it.
            if (isTable(next) || !next.style.keepWithNext || next.style.pageBreakBefore) {
                break;
            }
            end++;
        }

        const follower = blocks[end + 1];
        const usable = undefined !== follower
            && !(isTable(follower) ? follower.pageBreakBefore : follower.style.pageBreakBefore);

        // The block that need only get its FIRST line onto the page.
        const tail = usable ? follower : blocks[end]!;
        // ...and the last one that has to be there whole.
        const lastWhole = usable ? end : end - 1;
        if (lastWhole < from) {
            // A lone paragraph with nothing to keep it with. Breaking the page
            // for that would strand it alone on one of its own.
            return { end, heightPx: null };
        }

        const room = contentHeight();
        let height = 0;
        for (let index = from; index <= lastWhole; index++) {
            const member = blocks[index] as Paragraph;
            const style = member.style;
            const { metrics } = flowParagraph(
                member,
                atY + height + style.spaceBeforePx,
            );
            height += style.spaceBeforePx + style.spaceAfterPx
                + metrics.reduce((total, box) => total + box.heightPx, 0);

            // Past a page's worth the answer is settled: the group fits nowhere,
            // and what is left to add can only make it taller. Measuring the
            // rest of a long chain to reach a conclusion already reached is what
            // would make this quadratic.
            if (height > room) {
                return { end, heightPx: height };
            }
        }

        return {
            end,
            heightPx: height + (isTable(tail)
                ? tail.spaceBeforePx + firstRowHeight(tail)
                : tail.style.spaceBeforePx + paragraphLineHeight(tail)),
        };
    };

    /**
     * Flow a table's rows onto pages.
     *
     * A row moves WHOLE. Word can split a row across a page boundary, and this
     * does not yet — so a table with very tall rows leaves more white space
     * here than in the .docx. It never loses content, and a row taller than a
     * page is placed anyway rather than looping forever.
     */
    const placeTable = (declared: Table, blockIndex: number): void => {
        if (declared.pageBreakBefore && !isEmpty()) {
            newPage();
        }
        cursorY += declared.spaceBeforePx;

        const { table, offsetPx } = fitTable(declared, contentWidth());
        const { measured, spans, heights } = measureTable(table, measureParagraph);

        const headers = measured.slice(0, headerRows(table).length);
        const headerHeight = headers.reduce(
            (sum, row, rowIndex) => sum + (heights[rowIndex] ?? row.heightPx), 0,
        );

        /**
         * The first row of the run a merge holds together.
         *
         * A boundary a merged cell crosses is not a place the table can break:
         * breaking there would draw the top of the cell on one page and leave
         * its foot hanging below the bottom of it.
         */
        const groupStart: number[] = measured.map((_, rowIndex) => rowIndex);
        spans.forEach((rowSpans, rowIndex) => {
            for (const span of rowSpans) {
                for (let inside = rowIndex + 1; inside < rowIndex + span; inside++) {
                    groupStart[inside] = Math.min(groupStart[inside] ?? inside, groupStart[rowIndex] ?? rowIndex);
                }
            }
        });
        // ...and, once every merge has had its say, the run's total height.
        const groupHeight = (start: number): number => {
            let total = 0;
            for (let index = start; index < measured.length; index++) {
                if (index > start && groupStart[index] !== start) {
                    break;
                }
                total += heights[index] ?? 0;
            }

            return total;
        };

        const emit = (row: typeof measured[number], rowIndex: number, repeated: boolean): void => {
            // A row cut in two carries its own height; only a whole one takes
            // the height the merges settled on.
            const rowHeight = row === measured[rowIndex]
                ? heights[rowIndex] ?? row.heightPx
                : row.heightPx;

            // A repeated header is the SAME row drawn again; its note belongs
            // to the page the row itself landed on and is not owed twice.
            if (!repeated) {
                for (const id of footnotesInRow(row)) {
                    if (!current().notes.includes(id)) {
                        current().notes.push(id);
                    }
                }
            }

            // Every horizontal rule of a table sits in a gap of its OWN
            // width, between the content above it and the content below —
            // outer rules and inner ones alike. Measured: a three-row table of
            // 11.5pt lines and one-point rules ran 38.5pt from the line above
            // it to the line below, which is 3 x 11.5 plus 4 x 1.
            cursorY += ruleAbove(table, rowIndex);

            current().rows.push(placeRow(
                row,
                table,
                { xPx: contentLeft() + offsetPx, yPx: cursorY },
                rowHeight,
                { blockIndex, rowIndex, repeated },
            ));
            cursorY += rowHeight;
            if (rowIndex === table.rows.length - 1) {
                cursorY += ruleBelow(table);
            }
        };

        const repeatHeaders = (rowIndex: number): void => {
            // Headers repeat on the continuation — but never in front of
            // themselves, and never where the page could not then hold the
            // START of the row they head. A header filling the page on its own
            // would push the table forward a row at a time and never finish.
            //
            // ONE LINE of that row, not the whole of it: the row is broken up
            // to fit, so the whole is the wrong measure. Weighing the whole is
            // what left a 150-line row with no header at all, where LibreOffice
            // repeated one above each of the three pages it spans.
            const under = Math.min(
                ...(measured[rowIndex]?.cells ?? []).map((cell) => cell.lines[0]?.heightPx ?? 0),
            );

            if (rowIndex >= headers.length
                && headerHeight + (Number.isFinite(under) ? under : 0) <= contentHeight()) {
                headers.forEach((header, index) => emit(header, index, true));
            }
        };

        measured.forEach((row, rowIndex) => {
            // Only at the head of a run a merge holds together: asking again
            // inside one would break exactly where the merge forbids it.
            let needs = groupStart[rowIndex] === rowIndex
                ? groupHeight(rowIndex)
                : 0;

            // A HEADER is never the last thing on a page, so it asks for room
            // for what it heads as well. How much depends on whether
            // that row can be broken: a splittable one can start here with a
            // single line under the header, while one that cannot be split
            // must fit whole or go over with the header above it.
            //
            // Capped at the page, since the row may be taller than one and the
            // header still has to be put somewhere; the `!isEmpty()` guard
            // below is what stops a fresh page from breaking on its own.
            if (rowIndex < headers.length) {
                const headed = measured[headers.length];
                const under = undefined === headed
                    ? 0
                    : (table.rows[headers.length]?.cantSplit
                        ? heights[headers.length] ?? headed.heightPx
                        : Math.min(...headed.cells.map(
                            (cell) => cell.lines[0]?.heightPx ?? 0)));

                needs = Math.min(
                    (heights[rowIndex] ?? 0) + under,
                    contentHeight(),
                );
            }

            // The room this row's own notes will want, which is not yet owed:
            // a row is measured against the page it would leave behind.
            const brought = footnotesInRow(row)
                .filter((id) => !current().notes.includes(id));
            const limit = pageFoot() - reservedFor(
                { notes: [...current().notes, ...brought], carriedPx: current().carriedPx },
                cursorY + needs,
            );

            /**
                 * Put this row down, breaking it across as many columns as it
                 * takes.
                 *
                 * Emitting what was left of a split row whole put a hundred and
                 * twenty-nine lines on one page and ran them off the bottom of
                 * it; a row of a hundred and fifty lines spans three
                 * pages, and LibreOffice fills each of them in turn, repeating
                 * the header above every one.
                 *
                 * `splitRow` answers null when it can gain nothing — every line
                 * fits, or none does — so this ends whatever the room.
                 */
            const spill = (first: typeof row): void => {
                let rest = first;
                for (;;) {
                    const room = pageFoot() - reservedFor(current()) - cursorY;
                    const next = rest.heightPx <= room ? null : splitRow(rest, room);
                    if (null === next) {
                        break;
                    }

                    emit(next.upper, rowIndex, false);
                    nextColumn();
                    // Nought, not the height of what is left: the remainder is
                    // split again to fit, so what has to fit HERE is the
                    // headers themselves.
                    repeatHeaders(rowIndex);
                    rest = next.lower;
                }
                emit(rest, rowIndex, false);
            };

            // Moving a row only helps if a column could HOLD it. One taller
            // than the page skipped from page to page after room it would
            // never find, stranding whatever stood above it; such a
            // row is put down where it stands and `spill` breaks it up.
            if (needs > 0 && needs <= contentHeight()
                && cursorY + needs > limit && !isEmpty()) {
                // Word fills the page and carries the rest of the row over,
                // unless the row says not to. Moving a tall row whole leaves
                // the foot of every page it lands under empty.
                const halves = table.rows[rowIndex]?.cantSplit
                    ? null
                    : splitRow(row, limit - cursorY);

                if (null !== halves) {
                    emit(halves.upper, rowIndex, false);
                    nextColumn();
                    repeatHeaders(rowIndex);
                    spill(halves.lower);

                    return;
                }

                nextColumn();
                repeatHeaders(rowIndex);
            }

            // A row that refuses to be split and cannot FIT asks for something
            // no page can give. Measured: LibreOffice broke a
            // 150-line `w:cantSplit` row across three pages, repeating the
            // header above each — so where the flag cannot be obeyed it is set
            // aside rather than run off the page.
            spill(row);
        });

        cursorY += table.spaceAfterPx;
    };

    /**
     * The space the paragraph above already spent below itself.
     *
     * Two paragraphs do not add their spacing up: the gap between them is the
     * LARGER of the first's `w:spacing/@w:after` and the second's `@w:before`.
     * Measured against LibreOffice with the two set differently in
     * both orders — 10 after then 20 before, and 20 after then 10 — and the gap
     * came out 20.00 both times, which only `max` explains. Adding them put
     * every spaced paragraph in every document further down its page than the
     * file asks, and the page count follows.
     */
    let spentBelowPx = 0;

    blocks.forEach((block, paragraphIndex) => {
        if (isTable(block)) {
            placeTable(block, paragraphIndex);
            spentBelowPx = 0;

            return;
        }

        const style = block.style;
        const available = contentWidth() - style.indentLeftPx - style.indentRightPx;

        // The page break and the space before it both move the cursor, and a
        // float is placed against the paragraph's TOP — so nothing about this
        // paragraph can be measured until they have happened.
        if (style.pageBreakBefore && !isEmpty()) {
            newPage();
        }

        // `keepNext` binds this paragraph to what follows it, so moving the
        // whole group to a fresh page is the only way to honour it.
        //
        // Decided ONCE, at the group's first paragraph. Asking again at each
        // member would walk a group that fits on no page down the document one
        // paragraph per page, each member breaking away from the one above it.
        if (style.keepWithNext && paragraphIndex > keepDecidedThrough) {
            const group = keepGroup(paragraphIndex, cursorY);
            keepDecidedThrough = group.end;

            // Never off a page that is already empty: there is nowhere better
            // to go, and a group taller than any page would break forever.
            if (null !== group.heightPx && !isEmpty()
                && cursorY + group.heightPx > bottom()) {
                nextColumn();
            }
        }

        // A paragraph with no border of its own ends the box the one before
        // it opened, and so does one carrying a DIFFERENT border; one carrying
        // the same border extends it. Which is why this is decided here, at the
        // paragraph boundary, and not at the end of every paragraph.
        //
        // The bottom rule's room is taken as the box closes rather than when
        // its last line is placed, because until the next paragraph is in hand
        // there is no knowing the box has ended.
        if (undefined === style.borders
            || null === openBorder
            || !sameBorders(openBorder.borders, style.borders)) {
            endBorderRun();
        } else if (undefined !== style.borders.insideH) {
            // The box goes on, with a rule across it. LibreOffice took the
            // rule's whole width AND its space out of the flow — an 11.5pt
            // step became 14.5 for a three-point rule and 20.5 once six points
            // of space were added — but drew the rule half a width below the
            // text ABOVE it either way, at 787.989 both times. The space is
            // all below the rule, not split round it, which a fixture stating
            // no space could not have shown.
            openBorder = {
                ...openBorder,
                innerYPx: [
                    ...openBorder.innerYPx,
                    cursorY + style.borders.insideH.widthPx / 2,
                ],
            };
            cursorY += roomFor(style.borders.insideH);
        }

        // A field in the BODY is answered by the page it lands ON, and every
        // break that could move it has now been taken — its own
        // `pageBreakBefore`, and the one a `keepNext` group above it forced.
        // Resolving before those read the page the paragraph was LEAVING, and
        // a three-page fixture printed 1, 1, 2.
        //
        // It must still happen before the paragraph is measured, because `9`
        // and `10` are not the same width — which is the same reason the
        // furniture path resolves early.
        const paragraph = hasField(block)
            ? resolveFields([block], {
                pageNumber: numberOf(Math.max(0, pages.length - 1)),
                pageCount: options.totalPages ?? pages.length,
                ...(undefined === options.pageNumberFormat
                    ? {}
                    : { pageNumberFormat: options.pageNumberFormat }),
            })[0] as Paragraph
            : block;

        cursorY += Math.max(0, style.spaceBeforePx - spentBelowPx);

        if (undefined !== style.borders && null === openBorder) {
            cursorY += roomFor(style.borders.top);
        }

        // A paragraph is broken at the width of ONE column, and the lines it
        // carries into the next are drawn at that width — off the right edge
        // of a narrower one, and off the paper. So where the next column is a
        // different width, a paragraph that will not fit whole is moved to it
        // entire and broken again there.
        //
        // Word re-breaks only what carries over, which fills the first column;
        // this leaves that column short. Every line is the width of the column
        // it is drawn in either way, which is the part that has to be true.
        const startY = cursorY;

        // Measured before the floats are placed, because `placeFloat` needs the
        // column and the floats have to be on the page before the lines that
        // wrap around them are broken.
        const column: LineBox = {
            leftPx: contentLeft() + style.indentLeftPx,
            widthPx: contentWidth() - style.indentLeftPx - style.indentRightPx,
        };
        for (const float of paragraph.floats ?? []) {
            current().floats.push(placeFloat(float, geometry, column, startY));
        }

        let flowed = flowParagraph(paragraph, startY);
        let { lines, lineHeight, metrics, boxOf } = flowed;
        // How much of the paragraph's text earlier columns already took, and
        // the width they were broken at. A paragraph is re-broken only when
        // the column it carries into is a DIFFERENT width; at the same width
        // its lines are already right, and re-breaking would be work for
        // nothing.
        // The marker belongs to the paragraph's first line ONCE. What carries
        // into another column is a fresh flow whose first line is not that one.
        let markerPending = undefined !== paragraph.marker;
        let consumed = 0;
        let brokenAt = contentWidth();
        let heights = metrics.map((box) => box.heightPx);

        // How many lines still fit on the page we are on.
        /**
         * How many lines still fit on the page we are on.
         *
         * A line carrying a reference brings its note's room DOWN with it, so
         * the page it is measured against is shorter for that line and for
         * every line after it. Without that, the last line of a page would take
         * its note past the foot of the paper.
         */
        const fitting = (from: number): number => {
            let count = 0;
            let y = cursorY;
            const notes = [...current().notes];
            for (let i = from; i < lines.length; i++) {
                const brought = footnotesOn(lines[i]!).filter((id) => !notes.includes(id));

                // A note too tall for any page takes everything below the line
                // that brought it: measured against LibreOffice, the body
                // stopped at the referencing line and the note filled the rest.
                const owed = { notes, carriedPx: current().carriedPx };
                if (notesHeight(owed) > contentHeight()) {
                    break;
                }

                // A bordered paragraph's LAST line brings its bottom rule down
                // with it, and the rule has to land inside the writing area.
                // Measured against LibreOffice: a box needing 11.5pt of line
                // and 13 of rule either side stayed on a page with 65 filler
                // lines and moved off one with 66 — which is the boundary the
                // bottom room decides, and nothing else does.
                // Every line, not just the paragraph's last: the box closes
                // at a page break too, so any line may be the one whose bottom
                // rule has to fit. Measured — a box needing 25pt below it took
                // two lines onto a page that would have held four without.
                const rulePx = roomFor(style.borders?.bottom);
                const ahead = { notes: [...notes, ...brought], carriedPx: current().carriedPx };
                if (y + heights[i]! + rulePx > pageFoot() - reservedFor(ahead, y + heights[i]!)) {
                    break;
                }
                notes.push(...brought);
                y += heights[i]!;
                count++;
            }

            return count;
        };

        let placed = 0;
        while (placed < lines.length) {
            if (contentWidth() !== brokenAt) {
                // Carried into a column of another width: what is left is
                // broken again at THIS one, and the lines already placed stay
                // where they were.
                consumed += lines[placed]!.startsAt;
                const rest = { ...paragraph, runs: runsFrom(paragraph.runs, consumed) };
                flowed = flowParagraph(rest, cursorY);
                lines = flowed.lines;
                lineHeight = flowed.lineHeight;
                metrics = flowed.metrics;
                boxOf = flowed.boxOf;
                heights = metrics.map((box) => box.heightPx);
                brokenAt = contentWidth();
                placed = 0;
                if (0 === lines.length) {
                    break;
                }
            }

            let count = fitting(placed);

            if (0 === count) {
                // Nothing fits here. Move on, unless the page is already empty —
                // in which case the paragraph is taller than a page and must be
                // placed anyway or this loops forever.
                if (!isEmpty()) {
                    nextColumn();
                    continue;
                }
                count = 1;
            }

            if ((style.widowControl ?? widowOrphan) && count < lines.length - placed) {
                count = applyWidowOrphan(count, placed, lines.length, style, contentHeight(), lineHeight);
                if (0 === count) {
                    nextColumn();
                    continue;
                }
            }

            for (let i = 0; i < count; i++) {
                const lineIndex = placed + i;
                // The page owes this line's note room from here on.
                for (const id of footnotesOn(lines[lineIndex]!)) {
                    if (!current().notes.includes(id)) {
                        current().notes.push(id);
                    }
                }
                const firstLine = 0 === lineIndex;
                // The first line is narrower by its own indent, so that is the
                // width its alignment is measured against — not the column's.
                const indent = firstLine ? style.indentFirstLinePx : 0;
                // The same box the line was BROKEN at, which means the same
                // top: the heights of the lines before it, added the same way.
                const box = boxOf(lineIndex, heights.slice(0, lineIndex)
                    .reduce((total, each) => total + each, 0));
                const { offsetPx, wordSpacingPx } = alignLine(
                    lines[lineIndex]!,
                    box.widthPx - indent,
                    style.alignment,
                    lineIndex === lines.length - 1,
                );

                if (firstLine) {
                    markerPending = markerPending && 0 === consumed;
                }

                current().lines.push({
                    ...(firstLine && markerPending && undefined !== paragraph.marker
                        ? {
                            marker: {
                                run: paragraph.marker.run,
                                xPx: contentLeft()
                                    + style.indentLeftPx
                                    - paragraph.marker.offsetPx,
                            },
                        }
                        : {}),
                    line: lines[lineIndex]!,
                    xPx: box.leftPx + indent + offsetPx,
                    yPx: cursorY,
                    heightPx: heights[lineIndex]!,
                    baselinePx: metrics[lineIndex]!.baselinePx,
                    wordSpacingPx,
                    paragraphIndex,
                    lineIndex,
                });
                cursorY += heights[lineIndex]!;
                if (undefined !== style.borders) {
                    extendBorder(style.borders, style, cursorY - heights[lineIndex]!, cursorY);
                }
            }

            placed += count;
            if (placed < lines.length) {
                nextColumn();
            }
        }

        cursorY += style.spaceAfterPx;
        spentBelowPx = style.spaceAfterPx;
    });

    /**
     * Put a header or footer on the page.
     *
     * `topOf` is given the measured height because the two edges are anchored
     * differently: a header hangs from its distance off the top and a footer
     * stands on its distance off the bottom, so only the footer's top depends
     * on how tall it turned out.
     */
    const place = (
        blocks: readonly Block[] | undefined,
        context: FieldContext,
        topOf: (heightPx: number) => number,
    ): PlacedFurniture | undefined => {
        if (undefined === blocks || 0 === blocks.length) {
            return undefined;
        }

        // Fields are resolved BEFORE measuring, because "9" and "10" are not
        // the same width: substituting after the line was broken would leave
        // the text a digit wider than the line it was fitted to.
        const stacked = stackBlocks(resolveFields(blocks, context), writingWidth);
        const topPx = topOf(stacked.heightPx);

        return {
            topPx,
            heightPx: stacked.heightPx,
            // Stacked from its own origin; this is where that origin sits on
            // the paper.
            lines: stacked.lines.map(
                (line) => translateLine(line, geometry.marginLeftPx, topPx)),
            // A header may hold a table of its own — a logo beside an address
            // is one — and dropping it would take the whole thing off the page.
            rows: stacked.rows.map((row) => translateRow(row, geometry.marginLeftPx, topPx)),
            paragraphBorders: stacked.borders.map((box) => ({
                ...box,
                leftPx: box.leftPx + geometry.marginLeftPx,
                rightPx: box.rightPx + geometry.marginLeftPx,
                topPx: box.topPx + topPx,
                bottomPx: box.bottomPx + topPx,
                innerYPx: box.innerYPx.map((value) => value + topPx),
            })),
        };
    };

    const pageCount = options.totalPages ?? pages.length;

    /**
     * A page's notes, stacked against the FOOT of it.
     *
     * Measured against LibreOffice: the last note's line sits on the bottom
     * margin and the block grows upward from there, with the rule four points
     * above the first of them. So the notes are laid out from the bottom, not
     * from wherever the body happened to stop.
     */
    const placeFootnotes = (
        sheet: Sheet,
        boxBottomPx: number,
        roomPx: number,
        pending: PlacedLine[],
    ): { placed: PlacedFootnotes | undefined; carried: PlacedLine[] } => {
        // What this page owes: whatever the last one could not print, then its
        // own notes in the order they were referenced.
        const owed = [...pending];
        for (const id of sheet.notes) {
            owed.push(...stackBlocks(footnotesById.get(id) ?? [], writingWidth).lines);
        }
        if (0 === owed.length) {
            return { placed: undefined, carried: [] };
        }

        // Only as many as the reserve was taken for. The rest go overleaf,
        // which is what makes a note taller than its page a CONTINUED note
        // rather than one drawn over the text above it.
        const separator = separatorHeightPx(sheet.notes[0]);
        let taken = 0;
        let heightPx = 0;
        while (taken < owed.length) {
            const next = heightPx + owed[taken]!.heightPx;
            if (next + separator > roomPx && taken > 0) {
                break;
            }
            heightPx = next;
            taken++;
        }

        const top = boxBottomPx - heightPx;
        let y = top;
        const lines = owed.slice(0, taken).map((line) => {
            const placed = translateLine({ ...line, yPx: 0 }, geometry.marginLeftPx, y);
            y += line.heightPx;

            return placed;
        });

        return {
            placed: {
                separatorYPx: top - FOOTNOTE_RULE_GAP_PX,
                separatorLeftPx: geometry.marginLeftPx,
                separatorWidthPx: FOOTNOTE_RULE_WIDTH_PX,
                lines,
            },
            carried: owed.slice(taken),
        };
    };

    // Whatever box the last paragraph left open ends where the text does.
    closeBorder();

    let pendingNotes: PlacedLine[] = [];

    const laid = pages.map((sheet, pageIndex) => {
        const { lines, rows, floats } = sheet;
        const pageNumber = numberOf(pageIndex);
        const context: FieldContext = {
            pageNumber,
            pageCount,
            ...(undefined === options.pageNumberFormat
                ? {}
                : { pageNumberFormat: options.pageNumberFormat }),
        };
        const header = place(
            options.headerFor?.(pageIndex, pageNumber),
            context,
            () => geometry.headerDistancePx ?? 0,
        );
        const footer = place(
            options.footerFor?.(pageIndex, pageNumber),
            context,
            (heightPx) => geometry.heightPx - (geometry.footerDistancePx ?? 0) - heightPx,
        );

        const room = boxOf(pageIndex, pageNumber);
        // Exactly what the text left. Anything else and the reserve the flow
        // worked to and the space the notes are given would disagree, which is
        // how a note comes to be drawn over the last line of the body.
        const usedPx = Math.max(
            room.topPx,
            ...lines.map((placed) => placed.yPx + placed.heightPx),
            ...rows.map((placed) => placed.yPx + placed.heightPx),
        );
        const footnoteRoom = Math.max(0, room.bottomPx - usedPx);
        const { placed: footnotes, carried } = placeFootnotes(
            sheet,
            room.bottomPx,
            footnoteRoom,
            pendingNotes,
        );
        pendingNotes = carried;

        return {
            lines,
            rows,
            floats,
            geometry,
            sectionIndex: 0,
            pageNumber,
            paragraphBorders: sheet.borders,
            lineNumbers: [],
            ...(undefined === header ? {} : { header }),
            ...(undefined === footer ? {} : { footer }),
            ...(undefined === footnotes ? {} : { footnotes }),
        };
    });

    // A note longer than every page it could reach still has to be printed:
    // pages of nothing but notes carry the rest of it rather than losing it.
    while (pendingNotes.length > 0) {
        const pageIndex = laid.length;
        const pageNumber = numberOf(pageIndex);
        const room = boxOf(pageIndex, pageNumber);
        const { placed, carried } = placeFootnotes(
            { lines: [], rows: [], floats: [], notes: [], carriedPx: 0, borders: [] },
            room.bottomPx,
            room.bottomPx - room.topPx,
            pendingNotes,
        );

        if (carried.length === pendingNotes.length) {
            // Nothing fitted, and another page would not help either.
            break;
        }
        pendingNotes = carried;
        laid.push({
            lines: [],
            rows: [],
            floats: [],
            geometry,
            sectionIndex: 0,
            pageNumber,
            paragraphBorders: [],
            lineNumbers: [],
            ...(undefined === placed ? {} : { footnotes: placed }),
        });
    }

    return laid;
}

/**
 * Lay out a document whose sections have different paper.
 *
 * Each section is flowed on its own, because a section break STARTS A NEW
 * PAGE by definition — there is no continuity of cursor to preserve across
 * one. What does have to be preserved is the block numbering: a placed line
 * reports its index in the WHOLE document, so a caller can map it back to the
 * block it came from without knowing which section that was.
 *
 * A `continuous` break does not start a page, and is folded into the section
 * before it by the reader rather than handled here.
 */
export function layoutSections(sections: readonly Section[], options: LayoutOptions = {}): Page[] {
    const pages = layoutRun(sections, options);

    // `NUMPAGES` cannot be answered until the document has been laid out, so a
    // document that asks for it is laid out TWICE: once to count the pages and
    // once to print the count.
    //
    // Two passes rather than a loop that might not settle, because the second
    // pass cannot change the count. Furniture is drawn into room the writing
    // area already reserved, so turning "1" into "10" never moves a page break.
    //
    // A single-section document never needs the second pass at all: the count
    // `layoutPages` used was its own page count, which for one section IS the
    // document's. That is the common case, and it costs nothing.
    // A single-section document never needs the second pass for its FURNITURE:
    // the count `layoutPages` used was its own page count, which for one
    // section IS the document's. Its BODY is another matter — a body field is
    // answered while the flow is running, when only the pages laid so far
    // exist, so `page 1 of 3` came out `1 of 1` on the first page of a
    // one-section document.
    const inBody = sections.some(
        (section) => section.blocks.some((block) => hasField(block, 'numPages')));

    if (undefined !== options.totalPages
        || (!inBody && (sections.length < 2 || !usesPageCount(sections, pages)))) {
        return pages;
    }

    return layoutRun(sections, { ...options, totalPages: pages.length });
}

/**
 * Whether anything this document actually draws asks for the page COUNT.
 *
 * Asked of the FINISHED pages rather than of the sections, because which header
 * a page draws depends on the page: a `NUMPAGES` living only in the even-page
 * footer of the second section still has to be answered, and a variant no page
 * ever uses must not cost a second pass.
 *
 * The BODY is asked by the CALLER rather than here, because it changes the
 * short circuit as well as the answer: a one-section document skips this
 * question entirely for its furniture, and must not for its text.
 */
function usesPageCount(sections: readonly Section[], pages: readonly Page[]): boolean {
    const within = new Map<number, number>();

    for (const page of pages) {
        const index = within.get(page.sectionIndex) ?? 0;
        within.set(page.sectionIndex, index + 1);

        const section = sections[page.sectionIndex];
        const furniture = [
            ...(section?.headerFor?.(index, page.pageNumber) ?? []),
            ...(section?.footerFor?.(index, page.pageNumber) ?? []),
        ];

        if (furniture.some((block) => hasField(block, 'numPages'))) {
            return true;
        }
    }

    return false;
}

function layoutRun(sections: readonly Section[], options: LayoutOptions): Page[] {
    const pages: Page[] = [];
    let offset = 0;

    sections.forEach((section, sectionIndex) => {
        // Page NUMBERS are one-based: the next page to be added would be
        // number `pages.length + 1`.
        const wanted = section.startsOn;
        const needsEven = 'evenPage' === wanted;
        const needsOdd = 'oddPage' === wanted;
        const nextIsEven = 0 === (pages.length + 1) % 2;

        if (0 !== sectionIndex && ((needsEven && !nextIsEven) || (needsOdd && nextIsEven))) {
            // The blank belongs to the section BEFORE the break: it is the
            // tail of that run of pages, and Word prints it on that paper.
            //
            // It carries NO header or footer. Verified against LibreOffice: the
            // page inserted to reach an odd-page section start comes out
            // entirely empty, while the pages either side of it keep theirs.
            const previous = sections[sectionIndex - 1]!;
            pages.push({
                lines: [],
                rows: [],
                floats: [],
                geometry: previous.geometry,
                paragraphBorders: [],
                lineNumbers: [],
                sectionIndex: sectionIndex - 1,
                pageNumber: pages.length + 1,
            });
        }

        const laid = layoutPages(section.blocks, section.geometry, {
            ...(undefined === section.footnotes ? {} : { footnotes: section.footnotes }),
            ...options,
            ...(undefined === section.contentBox ? {} : { contentBox: section.contentBox }),
            ...(undefined === section.pageNumberFormat
                ? {}
                : { pageNumberFormat: section.pageNumberFormat }),
            ...(undefined === section.headerFor ? {} : { headerFor: section.headerFor }),
            ...(undefined === section.footerFor ? {} : { footerFor: section.footerFor }),
            // Where this section's first page lands in the WHOLE document,
            // which is what an alternating header counts.
            firstPageNumber: section.firstPageNumber ?? pages.length + 1,
        });

        for (const page of laid) {
            pages.push({
                ...page,
                lines: page.lines.map((line) => ({
                    ...line,
                    paragraphIndex: line.paragraphIndex + offset,
                })),
                rows: page.rows.map((row) => ({ ...row, blockIndex: row.blockIndex + offset })),
                geometry: section.geometry,
                // Resolved per page rather than once per section, because a
                // section's own paper is what it is measured against — and the
                // blank page inserted for an odd-page start carries none, the
                // same as it carries no header.
                ...(undefined === section.pageBorders
                    ? {}
                    : { pageBorder: placePageBorder(section.pageBorders, section.geometry) }),
                sectionIndex,
            });
        }

        offset += section.blocks.length;
    });

    return numberLines(pages, sections);
}

/**
 * Stack blocks into a column of unbounded height and say how tall they came
 * out — what a header, a footer or a table cell is.
 *
 * Tables are skipped: a header containing one is rare, and guessing its
 * height would move the body text by an amount nobody could account for.
 */
export function stackBlocks(
    blocks: readonly Block[],
    widthPx: number,
): StackedBlocks {
    return stackBlocksIn(blocks, widthPx, measureParagraph);
}

/**
 * Move a placed line into another coordinate space.
 *
 * The marker moves WITH it. Spreading the line and overriding only `xPx` leaves
 * the bullet behind in the space the line was stacked in — which is what both
 * of the callers used to do, and it put the marker a whole page margin away
 * from its text inside a table cell.
 */
/**
 * A paragraph's lines and the height of each, at a given column width.
 *
 * Shared by the page flow and by table cells so that the same text breaks the
 * same way inside a table as outside one. Two copies would differ eventually,
 * and the difference would look like a table bug.
 */
function measureParagraph(
    paragraph: Paragraph,
    availableWidth: number,
    boxOf?: (lineIndex: number, topPx: number) => LineBox,
): { lines: Line[]; lineHeight: number; metrics: readonly LineMetrics[] } {
    // A paragraph with NO runs draws nothing — not even an empty line. The
    // reader makes one only for a piece of a paragraph a page break split off
    // with nothing in it, which LibreOffice gives no room on either side of the
    // break. It still carries its `pageBreakBefore`, which is the
    // whole reason it exists rather than being dropped.
    //
    // An ordinary empty paragraph is NOT this: the reader gives it a run off
    // its paragraph mark, so it arrives here with one and takes its line.
    if (0 === paragraph.runs.length) {
        return { lines: [], lineHeight: 0, metrics: [] };
    }

    const style = paragraph.style;
    // Where the line about to be broken STARTS, measured from the paragraph's
    // own top. The lines before it are already broken, so their heights are
    // known exactly — counting them and multiplying by a nominal height is what
    // walked a float's band at the wrong speed.
    // None of them is the paragraph's LAST — there is always another coming,
    // which is the one being broken — so none gives back what the last does.
    const topOf = (linesSoFar: readonly Line[]): number => linesSoFar
        .reduce((total, line) => total + lineMetrics(line, paragraph, false).heightPx, 0);
    const widthOf = (lineIndex: number, linesSoFar: readonly Line[]): number =>
        boxOf?.(lineIndex, topOf(linesSoFar)).widthPx ?? availableWidth;
    const lines = breakIntoLines(
        paragraph.runs,
        (lineIndex, linesSoFar) => 0 === lineIndex
            ? widthOf(lineIndex, linesSoFar) - style.indentFirstLinePx
            : widthOf(lineIndex, linesSoFar),
        {
            tabStops: {
                ...(undefined === style.tabStops ? {} : { stops: style.tabStops }),
                ...(undefined === style.decimalSymbol
                    ? {}
                    : { decimalSymbol: style.decimalSymbol }),
                defaultPx: style.defaultTabPx ?? DEFAULT_TAB_PX,
                // Stops are measured from the MARGIN — the left edge of the
                // column this paragraph sits in — and NOT from the paragraph's
                // own indent (a stop at 80pt under a 36pt indent
                // printed at 80pt from the margin, not 116pt). So the origin
                // is everything between that edge and where this line starts:
                // the indent, plus the first line's own indent on line zero.
                //
                // Getting this wrong shifted every tab in an indented
                // paragraph by the indent, which for a numbered heading or an
                // indented contents line is the whole column.
                originOf: (lineIndex: number): number =>
                    style.indentLeftPx + (0 === lineIndex ? style.indentFirstLinePx : 0),
            },
        },
    );

    return {
        lines,
        lineHeight: paragraphLineHeight(paragraph),
        metrics: lines.map((line, index) => lineMetrics(line, paragraph, index === lines.length - 1)),
    };
}

/**
 * The box ONE line got, and where its baseline sits inside it.
 *
 * Measured against LibreOffice: a paragraph whose first line carries
 * a 36pt picture printed that line 38.10 tall and every line after it 11.50,
 * and moving the picture into the middle of the same paragraph moved the tall
 * line with it. So a picture makes ITS line taller and no other — the box
 * belongs to the line, not to the paragraph that owns it.
 *
 * A paragraph that states its own `w:spacing` is a FLOOR and not a ceiling —
 * see {@link boxHeightPx}, which is where that was measured.
 */
function lineMetrics(line: Line, paragraph: Paragraph, isLast: boolean): LineMetrics {
    // An empty line has nothing to measure but is still a line — its height is
    // the paragraph mark's own, which is what the runs say.
    const natural = naturalBox(0 === line.pieces.length ? paragraph.runs : line.pieces);
    const rule = paragraph.style.lineRule ?? 'auto';
    const declaredPx = paragraph.style.lineHeightPx;
    const heightPx = boxHeightPx(natural, paragraph, declaredPx, rule)
        - (isLast ? paragraphEndTrimPx(natural, paragraph, rule) : 0);

    return {
        heightPx,
        baselinePx: baselineOffsetPx({
            // A grid measures its leading against the PITCH, which the box may
            // be a multiple of — a 1.5-spaced line on an 18pt grid is 27.00
            // deep and still has its natural line centred in 18.
            lineHeightPx: 'grid' === rule
                ? paragraph.style.gridPitchPx ?? declaredPx ?? heightPx
                : heightPx,
            naturalHeightPx: natural.heightPx,
            // How deep the BOX goes below the baseline, which is the font's
            // descender except where a turned picture has moved the baseline
            // inside a box the picture set.
            descentPx: natural.belowPx,
            // ON a grid, every rule centres its natural line in a unit — the
            // pitch for the rest, its own declared height for `exact`, which
            // printed 15.62 down a 24pt line where the flat ratio this engine
            // uses off a grid gives 19.20.
            rule: 'exact' === rule && undefined !== paragraph.style.gridPitchPx ? 'grid' : rule,
        }),
    };
}

/**
 * How tall the line's box is: what the paragraph asked for against what the
 * line itself needs.
 *
 * ## Only `exact` is a ceiling
 *
 * Measured against LibreOffice with the same 18x36 picture under four
 * spacings. `exact` at 12pt kept its 12.00 and let the picture overflow into
 * the paragraph above — that is what "exactly" means, and the only case where
 * the declared height wins outright. `atLeast` at 12pt grew to 38.10, the
 * picture's own box, with the baseline 36.00 down and the picture standing on
 * it: a floor is a floor.
 *
 * ## `auto` ADDS its leading rather than replacing the line
 *
 * A proportional height is a multiple of the FONT's line and of nothing else,
 * so what it contributes is the leading — the same number on every line of the
 * paragraph, added to whatever that line happens to hold. At 1.5 over a 11.50pt
 * font that is 5.75, and the picture's line printed 43.85 rather than the 38.10
 * it takes unspaced, while the plain lines beside it printed 17.25.
 *
 * The same arithmetic explains a paragraph of MIXED sizes, which is what pinned
 * it down: at 1.5 over a paragraph whose tallest run is 23.00, the leading is
 * 11.50, and its 10pt first line printed 23.00 deep where the line carrying the
 * 20pt word printed 34.50. Neither is a multiple of its own line — both are
 * their own line plus the paragraph's leading. Subtracting rather than clamping
 * is what keeps spacing below single spacing compressing, which is its purpose.
 *
 * ## A grid line takes whole grid lines
 *
 * On an 18pt `w:docGrid`, the same picture's line printed 54.00 deep — three
 * pitches — with its baseline 36.00 from the top, so the line sits flush with
 * the top of the group and the spare room falls below it. A grid keeps its
 * rhythm by spending several of its own lines rather than by growing one.
 */
function boxHeightPx(
    natural: NaturalBox,
    paragraph: Paragraph,
    declaredPx: number | undefined,
    rule: LineRule,
): number {
    if (undefined === declaredPx || declaredPx <= 0) {
        return natural.heightPx;
    }

    if ('exact' === rule) {
        return declaredPx;
    }

    if ('grid' === rule) {
        // The room the paragraph asked of the grid, against the whole pitches
        // its own text needs — a 1.5-spaced line asks 27.00 of an 18pt grid and
        // a line carrying a 36pt picture needs 54.00, and the line takes
        // whichever is larger.
        const pitchPx = paragraph.style.gridPitchPx ?? declaredPx;

        return Math.max(declaredPx, Math.ceil(natural.heightPx / pitchPx) * pitchPx);
    }

    if ('atLeast' === rule) {
        return Math.max(declaredPx, natural.heightPx);
    }

    // The leading the multiple asked for, which the reader worked out against
    // the paragraph's tallest FONT — so taking that back off leaves the leading
    // alone, whatever this line turned out to hold.
    return Math.max(0, natural.heightPx + declaredPx - naturalBox(paragraph.runs).textPx);
}

/**
 * What the LAST line of a gridded paragraph gives back.
 *
 * A paragraph spaced above the grid's own pitch does not simply take that room
 * for every line: measured against LibreOffice on an 18pt grid over
 * an 11.50pt font, paragraphs at 1.5 lines totalled 23.75, 77.75 and 131.75 at
 * one, three and five lines — `N x 27.00` less a constant 3.25 — and at 2.0
 * lines the constant was 6.50 instead. That is `(multiple - 1)` times the
 * grid's own leading, spent once at the foot of the paragraph, and it is
 * nothing at all where the paragraph asks for no more than the pitch: a plain
 * grid paragraph and an `atLeast` one both totalled exactly `N x 18.00`.
 *
 * What LibreOffice is doing with that room is not clear from six numbers, and
 * this does not pretend otherwise — it reproduces the measurement rather than
 * explaining it. The `exact` rule is untouched, having no multiple to speak of.
 */
function paragraphEndTrimPx(natural: NaturalBox, paragraph: Paragraph, rule: LineRule): number {
    const pitchPx = paragraph.style.gridPitchPx;
    const declaredPx = paragraph.style.lineHeightPx;
    if ('grid' !== rule || undefined === pitchPx || undefined === declaredPx || pitchPx <= 0
        || true === paragraph.style.gridFloor) {
        return 0;
    }

    return Math.max(0, declaredPx / pitchPx - 1) * Math.max(0, pitchPx - natural.textPx);
}

/**
 * How tall a line of this paragraph is NOMINALLY — before the breaking says
 * which of its runs landed on which line.
 *
 * Separate from the measuring because a float has to know it BEFORE the lines
 * are broken: which lines a float displaces depends on where they sit, and
 * where they sit depends on their height. It does not depend on the breaking,
 * so there is no circle — only an order.
 *
 * A paragraph carrying a picture is taller here than most of its lines will
 * turn out to be. That over-estimates which lines a float displaces, and is
 * left alone deliberately: a float and an inline picture in one paragraph is a
 * case no measurement covers, and guessing at it would be inventing.
 */
function paragraphLineHeight(paragraph: Paragraph): number {
    return paragraph.style.lineHeightPx ?? naturalBox(paragraph.runs).heightPx;
}

/**
 * Give a float a place on the page.
 *
 * ## The offset is from an ORIGIN the document names
 *
 * `column` and `margin` both mean the writing area here — they differ only in
 * a multi-column layout, which this engine does not do. `page` is the paper's
 * own corner. Vertically, `paragraph` means the top of the paragraph the float
 * is anchored to, which is the common case and the only one whose origin is
 * not a fixed page coordinate.
 *
 * An `align` of left/right/centre is honoured against the same origin. Word
 * also allows `inside`/`outside`, which alternate with the page's parity; they
 * are read as left and right, which is right for a one-sided document.
 */
function placeFloat(
    float: Float,
    geometry: PageGeometry,
    column: LineBox,
    paragraphTopPx: number,
): PlacedFloat {
    const width = isFloatingBox(float) ? float.widthPx : float.image.widthPx;
    const height = isFloatingBox(float) ? float.heightPx : float.image.heightPx;

    const horizontalOrigin = 'page' === float.horizontal.relativeTo ? 0 : column.leftPx;
    const room = 'page' === float.horizontal.relativeTo ? geometry.widthPx : column.widthPx;

    let xPx = horizontalOrigin + float.horizontal.offsetPx;
    if (undefined !== float.horizontal.align) {
        const align = float.horizontal.align;
        const offset = 'right' === align || 'outside' === align
            ? room - width
            : ('center' === align ? (room - width) / 2 : 0);
        xPx = horizontalOrigin + offset;
    }

    const verticalOrigin = 'paragraph' === float.vertical.relativeTo
        ? paragraphTopPx
        : ('page' === float.vertical.relativeTo ? 0 : geometry.marginTopPx);
    const yPx = verticalOrigin + float.vertical.offsetPx;

    return {
        ...floatContent(float, xPx, yPx, width),
        xPx,
        yPx,
        widthPx: width,
        heightPx: height,
        wrap: float.wrap,
        behindText: float.behindText,
        // The clear space is part of the obstacle, not of the picture: text
        // keeps away from it, and nothing is drawn there.
        exclusion: {
            xPx: xPx - float.marginLeftPx,
            yPx: yPx - float.marginTopPx,
            widthPx: width + float.marginLeftPx + float.marginRightPx,
            heightPx: height + float.marginTopPx + float.marginBottomPx,
        },
    };
}

/**
 * What a placed float carries: the picture, or the box's own lines.
 *
 * A box is laid out HERE, at the moment it is placed, because that is the
 * first point at which its width is a number rather than a promise — and the
 * lines come back already at page coordinates, so nothing downstream has to
 * know a float can hold text at all.
 *
 * ## The inset is inside the frame, on all four sides
 *
 * Measured against LibreOffice: a 120pt DrawingML box silent on
 * `bodyPr` started its text 7.20 in from the frame and its first baseline
 * 3.60 below the top, and the same box stating `lIns="0" tIns="0"` started at
 * the frame's own edge on both axes. Those are the 0.1in and 0.05in the format
 * defines, so the file is believed and the defaults come from the spec.
 *
 * The lines then step exactly as they would outside a box — 11.50 apart for
 * the 10pt text that was probed, the body's own step — so a box needs nothing
 * of its own beyond the width it breaks at.
 */
function floatContent(
    float: Float,
    xPx: number,
    yPx: number,
    widthPx: number,
): { image: InlineImage }
    | { lines: readonly PlacedLine[]; rows: readonly PlacedRow[] } {
    if (!isFloatingBox(float)) {
        return { image: float.image };
    }

    const inner = Math.max(0, widthPx - float.inset.leftPx - float.inset.rightPx);
    const stacked = stackBlocks(float.blocks, inner);

    const dxPx = xPx + float.inset.leftPx;
    const dyPx = yPx + float.inset.topPx;

    return {
        lines: stacked.lines.map((line) => translateLine(line, dxPx, dyPx)),
        // A TABLE in the box travels the same way its paragraphs do;
        // `stackBlocks` has been laying them out all along.
        rows: stacked.rows.map((row) => translateRow(row, dxPx, dyPx)),
    };
}

/**
 * How far the deepest run on the line reaches below the baseline.
 *
 * The deepest rather than the first, for the same reason the line takes its
 * height from the tallest: one run with a longer descender decides where the
 * line's bottom is, and taking another run's would clip it.
 */
function deepestDescent(items: readonly BoxItem[]): number {
    let descent = 0;
    for (const item of items) {
        descent = Math.max(descent, item.font.descent(item.sizePx));
    }

    return descent;
}

/**
 * Number the body lines of every page, in place.
 *
 * After the pages are laid rather than during, because the count depends on
 * where the page breaks fell — and `restart: continuous` depends on every page
 * before this one, which a single page cannot see.
 *
 * `page.lines` is the body alone: a table's text lives in `page.rows` and a
 * note's in `page.footnotes`, so both are left unnumbered without a word — and
 * that is what LibreOffice printed, a table's line passing over silently while
 * an empty paragraph took a number of its own.
 */
function numberLines(pages: readonly Page[], sections: readonly Section[]): Page[] {
    let count = 0;
    let previousSection = -1;

    return pages.map((page) => {
        const numbering = sections[page.sectionIndex]?.lineNumbering;

        if ('newPage' === numbering?.restart
            || ('newSection' === numbering?.restart && page.sectionIndex !== previousSection)) {
            count = 0;
        }
        previousSection = page.sectionIndex;

        if (undefined === numbering) {
            return page;
        }

        const rightPx = page.geometry.marginLeftPx - numbering.distancePx;
        const numbers: PlacedLineNumber[] = [];

        // An EMPTY paragraph is a line and takes a number — measured, it took
        // the third of five — but it holds no piece to take a font from. The
        // page's own text stands in, which is the same font in every document
        // that does not change it mid-page.
        const fallback = page.lines
            .map((line) => line.line.pieces[0])
            .find((piece) => undefined !== piece);

        for (const line of page.lines) {
            // A paragraph may take itself out of the numbering, and it is not
            // COUNTED rather than counted and left blank: measured,
            // the paragraph above printed 10, the quiet one printed nothing,
            // and the next took 11.
            const source = sections[page.sectionIndex]?.blocks[line.paragraphIndex];

            if (undefined !== source && !isTable(source)
                && true === source.style.suppressLineNumbers) {
                continue;
            }
            count++;
            if (0 !== (numbering.start + count) % numbering.countBy) {
                continue;
            }

            // Drawn in the font of the line it stands beside, which is what
            // LibreOffice used — the digits came out the body's own size.
            const piece = line.line.pieces[0] ?? fallback;
            if (undefined === piece) {
                continue;
            }

            // A line number is furniture, not part of the run beside it, so it
            // carries no `w:kern` of its own and is measured unkerned like the
            // rest of a document that says nothing.
            const run: StyledRun = { text: String(numbering.start + count), font: piece.font, sizePx: piece.sizePx };
            const widthPx = advanceOf(run.text, run);

            numbers.push({
                run,
                // Right-aligned: a two-digit number reaches further left, which
                // is how 6 and 8 stood at 4.90 and 10 at −0.65.
                xPx: rightPx - widthPx,
                baselinePx: line.yPx + line.baselinePx,
            });
        }

        return { ...page, lineNumbers: numbers };
    });
}

/**
 * Where a page's own border is drawn, in the page's coordinates.
 *
 * The `w:space` is the clear gap between the rule and what it is measured
 * from, so the rule's CENTRE is half a width further on: a three-point border
 * 24 points from the paper's edge drew at 25.5, and the same border measured
 * from a 28.35pt margin drew at 2.85 — outside the text by the same 24.
 */
export function placePageBorder(
    declared: PageBorders,
    geometry: PageGeometry,
): PlacedParagraphBorder {
    const { borders, offsetFrom } = declared;
    const gapOf = (side: BorderSide | undefined): number =>
        undefined === side ? 0 : (side.spacePx ?? 0) + side.widthPx / 2;

    const inward = 'page' === offsetFrom;

    return {
        borders,
        leftPx: inward
            ? gapOf(borders.left)
            : geometry.marginLeftPx - gapOf(borders.left),
        rightPx: inward
            ? geometry.widthPx - gapOf(borders.right)
            : geometry.widthPx - geometry.marginRightPx + gapOf(borders.right),
        topPx: inward
            ? gapOf(borders.top)
            : geometry.marginTopPx - gapOf(borders.top),
        bottomPx: inward
            ? geometry.heightPx - gapOf(borders.bottom)
            : geometry.heightPx - geometry.marginBottomPx + gapOf(borders.bottom),
        opensHere: true,
        closesHere: true,
        innerYPx: [],
    };
}

/** The notes a line's marks stand for, in the order they are drawn. */
function footnotesOn(line: Line): number[] {
    const ids: number[] = [];
    for (const piece of line.pieces) {
        if (undefined !== piece.footnoteId && !ids.includes(piece.footnoteId)) {
            ids.push(piece.footnoteId);
        }
    }

    return ids;
}

/**
 * The notes referenced from anywhere in a row.
 *
 * A cell is a little page of its own, so a reference can be as deep inside one
 * as the document nests tables. Without this the MARK is drawn and the note it
 * points at never reaches a page — a reference to nothing, which is worse than
 * no reference at all.
 */
function footnotesInRow(row: { cells: readonly {
    lines: readonly PlacedLine[];
    rows: readonly PlacedRow[];
}[] }): number[] {
    const ids: number[] = [];
    const add = (found: readonly number[]): void => {
        for (const id of found) {
            if (!ids.includes(id)) {
                ids.push(id);
            }
        }
    };

    for (const cell of row.cells) {
        for (const line of cell.lines) {
            add(footnotesOn(line.line));
        }
        for (const nested of cell.rows) {
            add(footnotesInRow(nested));
        }
    }

    return ids;
}

/**
 * Adjust a split so it strands no single line.
 *
 * Returns how many lines to place on the current page, or 0 meaning "none of
 * it — start the paragraph on the next page".
 */
function applyWidowOrphan(
    count: number,
    placed: number,
    total: number,
    style: ParagraphStyle,
    contentHeight: number,
    lineHeight: number,
): number {
    const remainingAfter = total - placed - count;

    // "Keep lines together" is absolute: either it all fits here or it all
    // moves. Except when it cannot fit on any page, where moving it achieves
    // nothing and only wastes a page.
    if (style.keepLinesTogether && total * lineHeight <= contentHeight) {
        return count >= total - placed ? count : 0;
    }

    // ORPHAN: one line of a fresh paragraph stranded at the foot of the page.
    if (0 === placed && 1 === count && total > 1) {
        return 0;
    }

    // WIDOW: exactly one line would be left over for the next page, so push a
    // second one across to keep it company.
    if (1 === remainingAfter && count > 1) {
        const kept = count - 1;

        // But pushing one across must not STRAND the first line where it
        // stood: a paragraph beginning here that cannot keep two lines moves
        // entirely, or the rule against widows makes an orphan.
        //
        // Measured with the control asked for outright, which is the
        // one case both renderers agree on: a page with room for two lines of
        // a three-line paragraph took none of them, and all three printed
        // overleaf. This engine kept the first and moved the other two.
        return 0 === placed && kept < 2 ? 0 : kept;
    }

    return count;
}

/**
 * A line's own box: how tall single spacing makes it, and how deep below the
 * baseline it reaches.
 *
 * Two numbers rather than one because a picture can move the baseline without
 * changing the height. Everything that stands ON the baseline — text, an
 * upright picture, a shape — keeps the box's foot one font descender below it,
 * so the two are the same question there and a single height answered both
 * until a TURNED picture arrived.
 */
interface NaturalBox {
    readonly heightPx: number;
    readonly belowPx: number;
    /**
     * What the TEXT alone asked for, before any picture was considered.
     *
     * Kept apart because a proportional line height is a multiple of the font's
     * line and of nothing else: the leading it adds is the same number on every
     * line of the paragraph, and it is added to whatever that line holds.
     */
    readonly textPx: number;
}

/**
 * What a line box needs to know about one thing standing on the line.
 *
 * A run and a line PIECE both answer it, which is the point: a paragraph's
 * nominal box is measured over its runs before the text is broken, and each
 * line's own box over the pieces that landed on it.
 */
interface BoxItem {
    readonly font: TrueTypeFont;
    readonly sizePx: number;
    readonly baselineShiftPx?: number;
    readonly border?: BorderSide;
    readonly image?: InlineImage;
    readonly shape?: InlineShape;
    readonly ruby?: RubyGloss;
}

/**
 * Single spacing for what stands on ONE line: how far the items reach above
 * the baseline, and how far below it their box goes.
 *
 * Tallest rather than first, because one larger word on a line makes the whole
 * line taller — a line whose height came from its first run would overlap the
 * line below wherever a bigger font appears mid-sentence.
 *
 * ## ⚠️ Each SIDE of the baseline is maxed on its own
 *
 * The two sides used to be one number: the tallest whole line among the runs,
 * with the deepest descender taken separately beside it. Where every run shares
 * a FACE the two rules cannot be told apart -- scaling one face to two sizes
 * scales its ascent and its descent together, so the tallest total is also the
 * deepest descender -- and every measurement this box rested on was taken that
 * way, one face at several sizes or a picture.
 *
 * Two faces separate them. MEASURED 2026-08-25 against LibreOffice at 22pt
 * (`tools/probe-docx-line-box.php`, eleven paragraphs, the baseline step INSIDE
 * each): a line holding Carlito and Liberation Mono stepped **27.60**, where
 * the tallest whole line is Carlito's **26.855** and Carlito's reach above the
 * baseline plus Liberation Mono's descender is **27.55**. Carlito has the
 * taller line AND the shallower descender, so the old rule lost the difference
 * entirely -- 0.75pt a line, one line in every thirty-six.
 *
 * A run's own LINE GAP stays with its ascent rather than being maxed on its
 * own: Carlito beside Liberation Serif stepped **26.85**, Carlito's own line,
 * though the serif carries a 0.94pt gap that a separate maximum would have
 * added. Liberation Mono beside the same serif stepped **27.20** against a
 * predicted 27.15 -- the serif's ascent AND its gap over the mono's descender.
 *
 * `baselineOffsetPx` already placed the baseline at `height - descent`, which
 * is this box's ABOVE, so nothing there had to move.
 */
function naturalBox(items: readonly BoxItem[]): NaturalBox {
    let abovePx = 0;
    for (const run of items) {
        // A run moved off the baseline takes the line with it: measured
        // against LibreOffice, an 11.5pt line holding a run raised 6.9 stepped
        // 18.4 — the whole raise on top of the line it left.
        // A boxed run keeps its room above and below as well: an 11.5pt line
        // holding one came out 13.5, 17.5 and 21.5 for room of 1, 3 and 5.
        // A `w:ruby` gloss is a second line ABOVE this one, and the line grows
        // by the whole of it: measured, a 5pt gloss over a 10pt base
        // printed a 17.25 line where the same text unglossed prints 11.50, and
        // the baseline sank from 9.38 to 15.13 — the gloss's own 5.75, all of
        // it above. So it behaves exactly like a run raised off the baseline,
        // which is the term beside it.
        // `naturalLineHeight - descent` is the run's ascender plus its line
        // gap: everything it asks for ABOVE the baseline. The three terms
        // beside it were measured into the line's total and are kept above,
        // which is where they already sat -- the descent below has never
        // included any of them.
        abovePx = Math.max(
            abovePx,
            run.font.naturalLineHeight(run.sizePx) - run.font.descent(run.sizePx)
                + Math.abs(run.baselineShiftPx ?? 0)
                + 2 * borderRoomOf(run.border)
                + (undefined === run.ruby ? 0 : run.ruby.font.naturalLineHeight(run.ruby.sizePx)),
        );
    }

    // An inline picture STANDS ON the baseline, so the line has to hold the
    // whole picture above the baseline and the font's descender below it.
    // Taking the picture's height alone would let its foot sink into the line
    // beneath by exactly one descender.
    const descentPx = deepestDescent(items);
    const textPx = abovePx + descentPx;
    // Where the text alone rests: its own line, less the descender under it.
    const ascentPx = abovePx;
    let heightPx = textPx;
    let belowPx = descentPx;

    for (const run of items) {
        if (undefined !== run.shape) {
            heightPx = Math.max(heightPx, run.shape.heightPx + descentPx);
        }
        if (undefined === run.image) {
            continue;
        }

        // How far above the baseline this picture's TOP reaches, which for one
        // standing on the baseline is its whole height. A TURNED line puts it
        // somewhere else, measured against LibreOffice: `centred` halves the
        // way between the ascender edge and the picture's foot, `stacked`
        // hangs the text's own line entirely below the picture.
        const reachPx = 'centred' === run.image.acrossLine
            ? (ascentPx + run.image.heightPx) / 2
            : 'stacked' === run.image.acrossLine
                ? run.image.heightPx + ascentPx
                : run.image.heightPx;
        // The box keeps one descender below the PICTURE's foot however the
        // picture sits in it — measured, a centred picture overhangs its own
        // line rather than making it taller. `stacked` is the exception,
        // because there the whole text line is below the picture.
        const boxPx = run.image.heightPx + descentPx
            + ('stacked' === run.image.acrossLine ? ascentPx : 0);

        heightPx = Math.max(heightPx, boxPx);
        belowPx = Math.max(belowPx, boxPx - reachPx);
    }

    return { heightPx: heightPx > 0 ? heightPx : 0, belowPx, textPx };
}
