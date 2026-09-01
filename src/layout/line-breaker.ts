import type { BorderSide } from './borders.js';
import { isZeroAdvance, type TrueTypeFont } from '../font/truetype-font.js';
import type { NumeralStyle } from '../text/numerals.js';
import type { InlineImage, InlineShape } from './image.js';
import {
    allowsBreakBetween,
    forbidsBreakBetween,
    isBreakableSpace,
    isMandatoryBreak,
    isSpace,
    SOFT_HYPHEN,
} from './break-opportunities.js';

/**
 * A field whose text is not known until the page is chosen.
 *
 * Word stores the value it last computed alongside the instruction, so a run
 * carrying one already HAS text — the number from whenever the document was
 * last opened. That cached text is kept as the run's own, both because it is
 * the right thing to show when a field cannot be resolved and because it is
 * what the file says.
 */
export type FieldKind = 'page' | 'numPages';

/**
 * `w:u/@w:val`, folded onto the rules this engine can draw.
 *
 * OOXML has seventeen — wavy, dot-dash, three kinds of thick — and they are
 * read as `single`, `double` or `dotted`/`dashed` where those exist. The
 * PRESENCE of a rule is what a reader sees; whether it waves is not.
 */
export type UnderlineStyle = 'single' | 'double' | 'dotted' | 'dashed';

export interface Underline {
    readonly style: UnderlineStyle;
    /** `#RRGGBB`, when the rule is a different colour from the text. */
    readonly colorHex?: string;
}

/** A rule through the text — one line or two. */
export type Strike = 'single' | 'double';

/**
 * `w:vertAlign` — whether a run sits off the baseline.
 *
 * `baseline` is the explicit "on the line", which a run needs in order to
 * CANCEL a superscript its style asked for — so it is a value rather than the
 * absence of one.
 */
export type VerticalAlign = 'baseline' | 'superscript' | 'subscript';

/**
 * A stretch of text sharing one font and size.
 *
 * Line breaking takes RUNS rather than a string because real paragraphs change
 * style mid-sentence — a bold word is narrower or wider than the same letters
 * around it, and a breaker that only handles uniform text has to be rewritten
 * the moment a document arrives.
 */
export interface StyledRun {
    readonly text: string;
    readonly font: TrueTypeFont;
    readonly sizePx: number;
    /**
     * `#RRGGBB`, when the document asked for one.
     *
     * Absent means the renderer's own default rather than black: a page drawn
     * white-on-dark wants the default to follow the page, and a run that
     * genuinely asked for black — including one that asked for `auto` — says so.
     */
    readonly colorHex?: string;
    /** `#RRGGBB` painted behind the text — a highlight or a shaded run. */
    readonly highlightHex?: string;
    /**
     * `w:spacing` — extra space after EVERY character of this run.
     *
     * Tracking, and it changes the WIDTH: a run whose spacing went unread
     * breaks its line in the wrong place and the page count follows. Negative
     * condenses.
     */
    readonly letterSpacingPx?: number;
    /**
     * `w:kern` — whether this run's PAIRS are kerned.
     *
     * Absent means not, which is the format's default and the page's
     * (`AVAVAVAVAV` printed 11.60pt wider without the element than with it,
     * and the wider one is the plain sum of advances). A run measured kerned
     * when the document never asked for it is too narrow by that much, which
     * is enough to move a line break and everything after it.
     */
    readonly kerned?: boolean;
    /**
     * `w:footnoteReference` — the note this mark stands for.
     *
     * The TEXT is the number the reader worked out; this is the identity behind
     * it, which is what tells the page flow whose note to make room for.
     */
    readonly footnoteId?: number;
    /**
     * A rule under the text, and its colour when that differs from the text's.
     *
     * `w:u` carries its own `w:color`, because a document may underline in one
     * colour and write in another.
     */
    readonly underline?: Underline;
    /** `w:strike` and `w:dstrike` — a rule THROUGH the text. */
    readonly strike?: Strike;
    /**
     * How far this run's baseline moves, NEGATIVE upwards.
     *
     * Superscript and subscript are not a smaller font on the same line: the
     * glyphs sit off the baseline, and only the shift says by how much. The
     * SIZE is already reduced by the time a run gets here, so measurement needs
     * nothing extra — which is the point of resolving it in the reader.
     */
    readonly baselineShiftPx?: number;
    /**
     * `w:bdr` — a box round this run, drawn once per LINE the run covers.
     *
     * Measured against LibreOffice: a run broken over three lines came out as
     * three complete boxes, each closed top and bottom, and the continuation
     * lines began `w:space` plus the rule's width in from the margin to leave
     * room for the left edge of their own. Read but NOT yet drawn.
     */
    readonly border?: BorderSide;
    /**
     * Set when this run is a field's RESULT rather than literal text, so the
     * layout can replace it once it knows which page it is drawing.
     */
    readonly field?: FieldKind;
    /**
     * `\* roman` and friends — how the field writes its number.
     *
     * Absent is arabic. Front matter numbered i, ii, iii is the reason this
     * exists: without it every such page prints a digit.
     */
    readonly fieldFormat?: NumeralStyle;
    /**
     * A picture standing in this run's place.
     *
     * Its `text` is then a single object-replacement character, so the run
     * still occupies one unit of source and a caret can still be put either
     * side of it. The WIDTH comes from the picture, not from measuring that
     * character, which has no glyph worth the name.
     */
    readonly image?: InlineImage;
    /**
     * A drawn box standing in this run's place.
     *
     * Like an inline picture in every way that matters to a line: it has a
     * size and no glyphs, so the text around it measures and breaks the same.
     */
    readonly shape?: InlineShape;
    /**
     * `w:ruby` — a smaller gloss set ABOVE this run, furigana and its like.
     *
     * On the run rather than beside it because it moves with the text it
     * annotates: it makes the line taller, it decides the run's advance where
     * it is the wider of the two, and it breaks nowhere.
     */
    readonly ruby?: RubyGloss;
}

/**
 * The gloss over a `w:ruby` base.
 *
 * Measured against LibreOffice: a 5pt gloss over a 10pt base made
 * the line 17.25 instead of 11.50 — the base's own line plus the GLOSS's
 * natural line, all of it above the base's baseline, which sank from 9.38 to
 * 15.13. The gloss then sits at the top of that room on a baseline of its own.
 */
export interface RubyGloss {
    readonly text: string;
    readonly font: TrueTypeFont;
    readonly sizePx: number;
    readonly colorHex?: string;
    /** As on the run it glosses — a gloss is drawn as a string too. */
    readonly kerned?: boolean;
}

/**
 * How wide a stretch of one run's text is, tracking included.
 *
 * Every measurement of a run's text goes through here, so a run that asks for
 * spacing is the same width wherever it is asked about — the breaker deciding
 * where a line ends and the renderer drawing it cannot reach different answers.
 *
 * Word's own model: the space follows EVERY character, the last one included.
 * Measured against LibreOffice at half a point, one, two and five, which
 * matched it exactly. LibreOffice drops the final gap when the value is
 * NEGATIVE, and will not condense past 1.65pt a character whatever the glyph's
 * own width — neither of which Word does, so a heavily condensed run is
 * narrower here than LibreOffice prints it.
 */
export function advanceOf(
    text: string,
    style: { font: TrueTypeFont; sizePx: number; letterSpacingPx?: number; kerned?: boolean },
): number {
    const spacing = style.letterSpacingPx ?? 0;
    // Kerned only where the run says `w:kern` — see {@link StyledRun.kerned}.
    const width = true === style.kerned
        ? style.font.measureAdvance(text, style.sizePx).widthPt
        : style.font.measureUnkerned(text, style.sizePx);

    // Counted over the characters that take room. A zero-width joiner is not a
    // letter, so a tracked run must not be widened by one — the font already
    // ignores it, and charging spacing for it here would put the phantom width
    // straight back for exactly the runs that set `w:spacing`.
    return width + (0 === spacing
        ? 0
        : [...text].filter((c) => !isZeroAdvance(c.codePointAt(0) ?? 0)).length * spacing);
}

/** How wide a gloss is, and nought where there is none. */
export function rubyWidthPx(ruby: RubyGloss | undefined): number {
    return undefined === ruby ? 0 : advanceOf(ruby.text, ruby);
}

/**
 * Whether `piece` begins a boxed run's span, given what stands before it.
 *
 * A run is cut into one piece per WORD, so the room its box needs is owed once
 * where the run opens on a line and once where it closes — not once a word.
 * Which is also why a run broken over three lines came out of LibreOffice as
 * three complete boxes: each line opens the span again.
 */
export function opensBorderSpan(
    piece: LinePiece | undefined,
    before: LinePiece | undefined,
): boolean {
    return undefined !== piece
        && (piece.borderRoomPx ?? 0) > 0
        && piece.runIndex !== before?.runIndex;
}

/**
 * The room a boxed run keeps clear on EVERY side: its space and its whole rule.
 *
 * The box is drawn half a rule outside the text, so its outer edge — what the
 * text around it has to clear — is a whole rule out. Measured against
 * LibreOffice at four points: a one-point rule at spaces of 0, 2 and 4 gave
 * advances of 19.30, 21.30 and 23.30, and a three-point rule at space 2 gave
 * the same 23.30 as a one-point rule at space 4.
 */
export function borderRoomOf(border: BorderSide | undefined): number {
    return undefined === border ? 0 : (border.spacePx ?? 0) + border.widthPx;
}

/**
 * The runs from `offset` on, in the same units a line's `startsAt` counts.
 *
 * What a paragraph has LEFT once part of it has been placed. A paragraph that
 * carries into a column of another width has to be broken again at that width,
 * and breaking the whole of it again would place the beginning twice.
 */
export function runsFrom(runs: readonly StyledRun[], offset: number): StyledRun[] {
    if (0 === offset) {
        return [...runs];
    }

    const remaining: StyledRun[] = [];
    let seen = 0;
    for (const run of runs) {
        const end = seen + run.text.length;
        if (end > offset) {
            remaining.push(offset > seen ? { ...run, text: run.text.slice(offset - seen) } : run);
        }
        seen = end;
    }

    return remaining;
}

/**
 * The width available to a line — a constant, or a function of which line it is.
 *
 * The function form exists for first-line indent, where line 0 is narrower than
 * the rest. It generalises to text wrapping beside a floated image without the
 * breaker needing to know what a float is.
 *
 * The lines already broken come with the question, because the answer can
 * depend on how tall they turned out: a float narrows the lines it sits BESIDE,
 * and which lines those are is decided by where the earlier ones ENDED rather
 * than by counting them.
 */
export type LineWidth = number
    | ((lineIndex: number, linesSoFar: readonly Line[]) => number);

/**
 * Where a tab goes.
 *
 * A tab is not a character with a width — it is an instruction to advance to
 * the next STOP. Measuring its glyph gives whatever the font happens to put
 * at .notdef, which is 8.11px in Carlito and 12.45px in Liberation Serif, and
 * neither has anything to do with where the text lands.
 */
/**
 * How a stop treats the text that follows it.
 *
 * A `left` stop is an advance and nothing more. The other three are not: where
 * the text lands depends on how WIDE it is, so the tab's own width cannot be
 * known until the text after it has been measured. That is the whole reason
 * this is not simply a column position.
 */
export type TabAlign = 'left' | 'center' | 'right' | 'decimal';

/**
 * `w:leader` — what fills the gap a tab opens up.
 *
 * Drawn as repeated GLYPHS rather than as a rule: LibreOffice emits a run of
 * full stops, and a line would neither share the text's colour nor sit on its
 * baseline. `heavy` is Word's thick underscore and is drawn as an ordinary one,
 * which is the right character at the wrong weight.
 */
export type TabLeader = 'dot' | 'hyphen' | 'underscore' | 'middleDot' | 'heavy';

export interface TabStop {
    readonly positionPx: number;
    readonly align: TabAlign;
    /** Absent for the usual blank tab. */
    readonly leader?: TabLeader;
}

export interface TabStops {
    /** Explicit stops, in ascending order, measured from the tab ORIGIN. */
    readonly stops?: readonly TabStop[];
    /**
     * What a `decimal` stop lines up on.
     *
     * `w:decimalSymbol` in settings.xml, which is a comma across most of
     * Europe. A document that says comma and is aligned on a full stop has its
     * numbers ranged on nothing at all.
     */
    readonly decimalSymbol?: string;
    /** The repeating stop past the last explicit one. Word's default is half an inch. */
    readonly defaultPx?: number;
    /**
     * How far line N starts from the tab origin.
     *
     * Stops belong to the COLUMN, not to the line and not to the paragraph:
     * they are measured from the margin, so an indented paragraph's tabs land
     * in the same places as an unindented one's. Every line therefore
     * reports how far along the stops its own start sits — its paragraph's
     * indent, plus the first line's own indent on line zero.
     */
    readonly originOf?: (lineIndex: number) => number;
}

export interface BreakOptions {
    readonly tabStops?: TabStops;
}

/** The part of one run that landed on one line. */
export interface LinePiece {
    readonly runIndex: number;
    readonly text: string;
    readonly widthPx: number;
    /**
     * The face and size this piece was MEASURED with.
     *
     * Carried on the piece rather than looked up through `runIndex`, because a
     * placed line travels without its runs: a renderer holding only a page
     * would otherwise have to reach back into the document to find out what to
     * draw the text in, and could reach a different answer than the measurement
     * did.
     */
    readonly font: TrueTypeFont;
    readonly sizePx: number;
    /** The run's colour, carried for the same reason the font is. */
    readonly colorHex?: string;
    /** The colour behind it, carried for the same reason again. */
    readonly highlightHex?: string;
    /**
     * `w:spacing` — extra space after EVERY character of this run.
     *
     * Tracking, and it changes the WIDTH: a run whose spacing went unread
     * breaks its line in the wrong place and the page count follows. Negative
     * condenses.
     */
    readonly letterSpacingPx?: number;
    /**
     * `w:kern`, carried for the same reason the font is — and for one more.
     *
     * A renderer draws a piece as a STRING and lets the drawing engine step
     * between the glyphs, and every drawing engine kerns by default. Told
     * nothing, it would kern text this engine measured unkerned, and the words
     * would not sit where the line breaker put them.
     */
    readonly kerned?: boolean;
    /**
     * `w:footnoteReference` — the note this mark stands for.
     *
     * The TEXT is the number the reader worked out; this is the identity behind
     * it, which is what tells the page flow whose note to make room for.
     */
    readonly footnoteId?: number;
    /** The rule under it, carried for the same reason again. */
    readonly underline?: Underline;
    /** The rule through it. */
    readonly strike?: Strike;
    /** How far this piece's baseline moves, negative upwards. */
    readonly baselineShiftPx?: number;
    /** `w:bdr`, carried from the run so a renderer can box this piece. */
    readonly border?: BorderSide;
    /**
     * What that box keeps clear on each side of the RUN this piece belongs to.
     *
     * Kept OUT of `widthPx`, which stays the text's own advance: the room is
     * the run's, not the word's, and is charged once where the run's span
     * opens on a line and once where it closes. See {@link opensBorderSpan}.
     */
    readonly borderRoomPx?: number;
    /** The picture this piece draws instead of text. */
    readonly image?: InlineImage;
    /**
     * A drawn box standing in this run's place.
     *
     * Like an inline picture in every way that matters to a line: it has a
     * size and no glyphs, so the text around it measures and breaks the same.
     */
    readonly shape?: InlineShape;
    /**
     * The gloss set above this piece, when it carries one.
     *
     * Its width may EXCEED the piece's own text, and then it is the piece's
     * width: measured, a gloss wider than its base pushed the text
     * after it along by the gloss and centred the base underneath.
     */
    readonly ruby?: RubyGloss;
    /**
     * The line may END here, drawing a hyphen, because the author offered it.
     *
     * `w:softHyphen`. Nothing is drawn where the line carries on past it, so
     * this is a property of the piece and the DECISION is the renderer's, which
     * is the only place that knows which piece ended the line.
     */
    readonly hyphenAtEnd?: boolean;
    /**
     * The fill this piece is a TAB for, when its stop asked for one.
     *
     * On the piece rather than worked out by a renderer because only the
     * breaker knows which stop the tab landed on — and the piece's own width is
     * exactly the span the fill has to cover.
     */
    readonly leader?: TabLeader;
    /**
     * Where this piece starts in the runs' text, concatenated — counted in
     * UTF-16 units, the same way a text offset is counted everywhere else.
     */
    readonly sourceStart: number;
}

export interface Line {
    readonly pieces: readonly LinePiece[];
    /**
     * Width used for fitting: trailing spaces are excluded, because a space at
     * a line end hangs into the margin rather than pushing the break earlier.
     * Getting this wrong breaks lines one word too soon on justified text.
     */
    readonly widthPx: number;
    /**
     * True when a single unbreakable stretch was wider than the column — a URL
     * or a long identifier. It is placed alone and allowed to overflow rather
     * than chopped at an arbitrary character, because a break the author did
     * not ask for is a change to their document.
     */
    readonly overflows: boolean;
    /** True when the author's own break ended this line, not the column width. */
    readonly endedByMandatoryBreak: boolean;
    /**
     * The span of source text this line covers, in UTF-16 units over the runs'
     * concatenated text.
     *
     * Summing the pieces' lengths does NOT give this: a mandatory break is a
     * position rather than ink, so it occupies source but appears in no piece.
     * An editor mapping a page boundary back to a caret position needs the
     * source span, and computing it from the pieces drifts by one character per
     * line break.
     */
    readonly startsAt: number;
    readonly endsAt: number;
    /**
     * How many inter-word gaps on this line can absorb slack when the
     * paragraph is justified.
     *
     * Counted HERE rather than by re-reading the pieces' text, because the
     * breaker is the only place that knows which spaces are interior: the run
     * of spaces that ends a line is trimmed off its width and stretching it
     * would push the last word past the margin.
     */
    readonly spaceGaps: number;
}

/**
 * Greedy line breaking — take the last opportunity that still fits.
 *
 * Greedy rather than optimal (Knuth-Plass) on purpose: Word and LibreOffice are
 * both greedy, and the engine's job is to agree with the .docx it produces, not
 * to set better type than it. An optimal breaker would look nicer on screen and
 * be WRONG about where the page ends.
 *
 *  Kerning is applied within a run but not across a run boundary. The pair
 * straddling a style change is one glyph from each font, and the fonts disagree
 * about what to do with it; renderers differ here too. The error is at most one
 * pair per style change and is documented rather than silently absorbed.
 */
export function breakIntoLines(
    runs: readonly StyledRun[],
    maxWidth: LineWidth,
    options: BreakOptions = {},
): Line[] {
    const widthOf = 'function' === typeof maxWidth ? maxWidth : (): number => maxWidth;
    const tabs = options.tabStops ?? {};
    const originOf = tabs.originOf ?? ((): number => 0);
    let origin = originOf(0);
    const segments = buildSegments(runs);
    const decimalSymbol = tabs.decimalSymbol ?? '.';
    const lines: Line[] = [];
    let maxWidthPx = widthOf(0, lines);

    let pieces: LinePiece[] = [];
    let widthWithSpaces = 0;
    let widthWithoutTrailing = 0;
    let overflows = false;
    let startsAt = 0;
    let endsAt = 0;
    let spacedSegments = 0;

    const flush = (mandatory: boolean): void => {
        lines.push({
            pieces,
            widthPx: widthWithoutTrailing,
            overflows,
            endedByMandatoryBreak: mandatory,
            startsAt,
            endsAt,
            // Every segment carries the spaces that follow it, so the number of
            // spaced segments IS the number of gaps — less one when the last of
            // them ends the line, since those spaces were trimmed off the width
            // and stretching them would move the margin, not the words.
            spaceGaps: Math.max(0, spacedSegments - (widthWithSpaces > widthWithoutTrailing ? 1 : 0)),
        });
        pieces = [];
        widthWithSpaces = 0;
        widthWithoutTrailing = 0;
        overflows = false;
        spacedSegments = 0;
        // The NEXT line may be a different width — a first-line indent narrows
        // line 0 only, and later a float will narrow whichever lines it sits
        // beside. Asking per line costs nothing and avoids a second breaker.
        maxWidthPx = widthOf(lines.length, lines);
        origin = originOf(lines.length);
    };

    for (let index = 0; index < segments.length; index++) {
        const segment = segments[index]!;
        const isEmptyLine = 0 === pieces.length;
        // A tab's width is the distance to its stop, which depends on where the
        // line has got to — so it can only be known here, not when the segment
        // was measured. For an aligned stop it depends on the text AFTER it as
        // well, which is why the following segments are measured here too.
        const advance = segment.isTab
            ? tabAdvance(segments, index, widthWithSpaces + origin, tabs, decimalSymbol)
            : segment.contentWidthPx;
        // A boxed run opening here owes its room at BOTH ends of its span,
        // and the far end is owed the moment it opens: a line that stops mid
        // span still closes its box, and has to have kept the room to do it.
        const roomFor = (): number =>
            opensBorderSpan(segment.pieces[0], pieces[pieces.length - 1])
                ? 2 * (segment.pieces[0]!.borderRoomPx ?? 0)
                : 0;
        const fits = widthWithSpaces + roomFor() + advance <= maxWidthPx + WIDTH_EPSILON;

        if (!isEmptyLine && !fits) {
            flush(false);
        }

        // Asked again AFTER the flush: a span carried onto the next line opens
        // there in its own right, and owes the room there too.
        const spanRoomPx = roomFor();

        // A line starts where its first segment starts and ends where its last
        // one ends. Because every character belongs to exactly one segment,
        // consecutive lines meet exactly — no gap for a position to fall into,
        // no overlap for it to fall into twice.
        if (0 === pieces.length) {
            startsAt = segment.sourceStart;
        }
        endsAt = segment.sourceEnd;

        // A word too long for an empty line is CHOPPED at the last character
        // that fits, and what is left of it starts the next line. Measured
        // against LibreOffice: `gggggggghhhhhhhh` in a 51.3pt measure
        // printed as `ggggggggh` and `hhhhhhh`, and the same word in 31.3pt as
        // five, five, five and one. A word of MIXED widths is what proves it is
        // fitted rather than counted — `WWWWiiiiWWWWiiii` broke after eight in
        // the measure that took nine of the even one.
        //
        // Chopping a word is not something anyone wants to read, and it is what
        // both Word and LibreOffice do rather than let it run off the page.
        if (0 === pieces.length && segment.contentWidthPx > maxWidthPx + WIDTH_EPSILON) {
            const split = splitToFit(segment, maxWidthPx - spanRoomPx);

            if (null === split) {
                // Not even one character fits, so there is nothing to chop
                // into: placing it whole is the only thing that terminates.
                overflows = true;
            } else {
                segments.splice(index, 1, split.head, split.tail);
                index--;
                continue;
            }
        }

        if (segment.isTab) {
            // The tab still occupies its source span and still has to be there
            // for an editor to map a position onto. It is drawn as space —
            // unless its stop asked for a leader, which fills the whole span.
            const leader = nextStop(widthWithSpaces + origin, tabs).leader;
            pieces.push(...segment.pieces.map((piece) => ({
                ...piece,
                widthPx: advance,
                ...(undefined === leader ? {} : { leader }),
            })));
            widthWithSpaces += spanRoomPx + advance;
            widthWithoutTrailing = widthWithSpaces;
        } else {
            pieces.push(...segment.pieces);
            widthWithSpaces += spanRoomPx + segment.fullWidthPx;
            widthWithoutTrailing = widthWithSpaces - segment.trailingSpaceWidthPx;
            if (segment.trailingSpaceWidthPx > 0) {
                spacedSegments++;
            }
        }

        if (segment.endsWithMandatoryBreak) {
            flush(true);
        }
    }

    if (0 !== pieces.length || 0 === lines.length) {
        flush(false);
    }

    return lines;
}

/**
 * One unbreakable stretch of text, plus whatever spaces trail it.
 *
 * Keeping the spaces WITH the preceding segment is what makes a wrapped line
 * start at the next word rather than at the space after the previous one.
 */
interface Segment {
    readonly pieces: LinePiece[];
    /** Everything, spaces included — the cost of putting the next word after it. */
    readonly fullWidthPx: number;
    /** Without the trailing spaces — the cost of ENDING a line here. */
    readonly contentWidthPx: number;
    readonly trailingSpaceWidthPx: number;
    readonly endsWithMandatoryBreak: boolean;
    /** True when this segment is a TAB, whose width depends on where it lands. */
    readonly isTab: boolean;
    /** Source span, INCLUDING any mandatory break that ends it. */
    readonly sourceStart: number;
    readonly sourceEnd: number;
}

/**
 * One segment cut at the last character that fits, and what is left of it.
 *
 * Null where not even one character fits: there is nothing to cut into then,
 * and cutting nothing off would put the same segment back for ever.
 *
 * A piece standing for a PICTURE has one character and no glyphs, so it is
 * never cut into — the cut falls before it or the picture goes on whole.
 */
function splitToFit(
    segment: Segment,
    maxWidthPx: number,
): { head: Segment; tail: Segment } | null {
    let usedPx = 0;

    for (const [index, piece] of segment.pieces.entries()) {
        // What ENDING here would cost on top of the piece itself: a piece that
        // finishes at the author's offer owes the hyphen that is then drawn.
        // Without this the whole piece looks as though it fits, the cut is
        // never looked for, and the line is placed with a hyphen hanging past
        // its own margin.
        const endCostPx = true === piece.hyphenAtEnd ? advanceOf('-', piece) : 0;

        if (usedPx + piece.widthPx + endCostPx <= maxWidthPx + WIDTH_EPSILON) {
            usedPx += piece.widthPx;
            continue;
        }

        const splittable = undefined === piece.image && undefined === piece.shape;
        const characters = splittable ? [...piece.text] : [];
        // Prefixes are measured WHOLE rather than summed a character at a time,
        // so the kerning of every pair inside the prefix is charged where it
        // falls — summing advances would let a word creep past its own line.
        let take = 0;
        let takePx = 0;
        for (let count = 1; count <= characters.length; count++) {
            // Cutting short of the author's offer forfeits it: the head keeps
            // no hyphen, and only the tail can still end at one.
            const px = advanceOf(characters.slice(0, count).join(''), piece)
                + (count === characters.length ? endCostPx : 0);
            if (usedPx + px > maxWidthPx + WIDTH_EPSILON) {
                break;
            }
            // A cut is still refused where the author refused one: `10 kg`
            // joined by U+00A0 is kept whole even by a line with no room.
            const next = characters[count];
            if (undefined !== next
                && forbidsBreakBetween(
                    characters[count - 1]!.codePointAt(0) ?? 0, next.codePointAt(0) ?? 0)) {
                continue;
            }
            take = count;
            takePx = px;
        }

        // TERMINATION lives here. The caller puts the head back through the
        // same loop, so the head must be strictly SHORTER than what came in —
        // a cut that keeps everything, or that returns a head as long as the
        // segment, feeds the same segment round for ever. Returning null when
        // not one character fits is what makes the head shorter every time,
        // and a mutation that breaks it does not fail the suite: it HANGS it,
        // which is why the batteries that test this file run under a timeout.
        if (0 === take && 0 === index) {
            return null;
        }

        const headText = characters.slice(0, take).join('');
        const tailText = characters.slice(take).join('');
        const headPieces = segment.pieces.slice(0, index);
        const tailPieces = segment.pieces.slice(index + 1);

        if ('' !== headText) {
            const { hyphenAtEnd: _forfeited, ...rest } = piece;

            headPieces.push({ ...rest, text: headText, widthPx: takePx });
        }
        if ('' === headText || '' !== tailText) {
            tailPieces.unshift({
                ...piece,
                ...('' === headText ? {} : {
                    text: tailText,
                    widthPx: advanceOf(tailText, piece),
                    sourceStart: piece.sourceStart + headText.length,
                }),
            });
        }

        const headWidthPx = headPieces.reduce((sum, each) => sum + each.widthPx, 0);
        const tailWidthPx = tailPieces.reduce((sum, each) => sum + each.widthPx, 0);
        const cutAt = segment.sourceStart
            + headPieces.reduce((sum, each) => sum + each.text.length, 0);

        return {
            head: {
                pieces: headPieces,
                fullWidthPx: headWidthPx,
                contentWidthPx: headWidthPx,
                trailingSpaceWidthPx: 0,
                endsWithMandatoryBreak: false,
                isTab: false,
                sourceStart: segment.sourceStart,
                sourceEnd: cutAt,
            },
            tail: {
                pieces: tailPieces,
                // The trailing spaces are already among the tail's own pieces,
                // so `full` is their sum and `content` is that sum without
                // them — adding the trailing width again would charge a line
                // for a space it does not have to fit, and chop one letter
                // early at the end of every wrapped word.
                fullWidthPx: tailWidthPx,
                // The offer travels with the tail, and so does what it costs to
                // accept: without this the tail rejoins the line the head just
                // left, and the hyphen is drawn for free past its margin.
                contentWidthPx: tailWidthPx - segment.trailingSpaceWidthPx
                    + (true === tailPieces[tailPieces.length - 1]?.hyphenAtEnd
                        ? advanceOf('-', tailPieces[tailPieces.length - 1]!)
                        : 0),
                trailingSpaceWidthPx: segment.trailingSpaceWidthPx,
                endsWithMandatoryBreak: segment.endsWithMandatoryBreak,
                isTab: false,
                sourceStart: cutAt,
                sourceEnd: segment.sourceEnd,
            },
        };
    }

    return null;
}

const TAB = 0x0009;

/**
 * How close two widths have to be to count as equal.
 *
 * Lengths reach this engine through twips and points, and the same length
 * computed two ways does not come back bit-identical: a writing area of
 * `page - left - right` and a tab stop declared at exactly that width differ by
 * about 1e-13px. Compared exactly, a right-aligned tab stop at the margin —
 * which is every table of contents ever written — pushes its page number onto a
 * line of its own.
 *
 * A ten-thousandth of a pixel is far below anything a device can show and far
 * above the noise. A tolerance, not a measurement of anything: no page was
 * printed to arrive at it.
 */
const WIDTH_EPSILON = 0.0001;

/**
 * How far a tab moves the line on.
 *
 * A LEFT stop simply advances to its column. The other three place the text
 * that FOLLOWS so it ends at the stop, straddles it, or lines its decimal
 * separator up with it — so the following text has to be measured before the
 * tab's own width can be known.
 *
 * ## A tab never moves backwards
 *
 * Text too wide to fit before an aligned stop leaves no room for the tab, and
 * the tab collapses to nothing rather than dragging the text back over what
 * came before it. Verified against LibreOffice, where a right-aligned tab
 * followed by forty-eight characters puts them immediately after the text
 * preceding the tab.
 */
function tabAdvance(
    segments: readonly Segment[],
    index: number,
    x: number,
    tabs: TabStops,
    decimalSymbol: string,
): number {
    const stop = nextStop(x, tabs);
    if ('left' === stop.align) {
        return stop.positionPx - x;
    }

    const width = alignedWidth(segments, index + 1, stop.align, decimalSymbol);
    const share = 'center' === stop.align ? width / 2 : width;

    return Math.max(0, stop.positionPx - share - x);
}

/**
 * How much of the text after a tab is measured against its stop.
 *
 * Everything up to the NEXT tab, because that is where the next stop takes
 * over. For a decimal stop it is only the part before the separator: what lines
 * up is the separator itself, and the digits after it hang past the stop.
 */
function alignedWidth(
    segments: readonly Segment[],
    from: number,
    align: TabAlign,
    decimalSymbol: string,
): number {
    let width = 0;

    for (let index = from; index < segments.length; index++) {
        const segment = segments[index]!;
        if (segment.isTab) {
            break;
        }

        if ('decimal' === align) {
            const upTo = decimalPrefixWidth(segment, decimalSymbol);
            if (null !== upTo) {
                return width + upTo;
            }
        }

        width += segment.contentWidthPx;
        if (segment.endsWithMandatoryBreak) {
            break;
        }
    }

    return width;
}

/**
 * The width of this segment up to its decimal separator, or null when it holds
 * none.
 *
 * Measured piece by piece so the separator is found in the run it actually sits
 * in: a number split across a style change still lines up on its point.
 */
function decimalPrefixWidth(segment: Segment, decimalSymbol: string): number | null {
    let width = 0;

    for (const piece of segment.pieces) {
        const at = piece.text.indexOf(decimalSymbol);
        if (at >= 0) {
            return width + advanceOf(piece.text.slice(0, at), piece);
        }
        width += piece.widthPx;
    }

    return null;
}

/**
 * The first stop past `x`.
 *
 * Explicit stops win while there are any; past the last one the default stops
 * take over, which is what makes a tabbed line keep advancing rather than stop
 * dead at the last declared column. Those repeating stops are always LEFT —
 * Word gives no way to align them.
 *
 * The defaults are the COLUMN's own, at multiples of the step from the margin,
 * and NOT a repeat starting from the last explicit stop (a stop at
 * 15pt with a 36pt default sent its tab to 36, where stepping from the stop
 * gives 51). The ones behind the last explicit stop are unreachable rather
 * than removed, which the loop above already sees to: anything reaching here
 * has passed every explicit stop there is.
 */
function nextStop(x: number, tabs: TabStops): TabStop {
    for (const stop of tabs.stops ?? []) {
        if (stop.positionPx > x + WIDTH_EPSILON) {
            return stop;
        }
    }

    const step = tabs.defaultPx ?? 0;
    if (step <= 0) {
        // No stops at all: a tab can still not be zero wide, or two tabbed
        // columns would print on top of each other.
        return { positionPx: x, align: 'left' };
    }

    return {
        positionPx: Math.floor(x / step + 1) * step,
        align: 'left',
    };
}

function buildSegments(runs: readonly StyledRun[]): Segment[] {
    const characters = flatten(runs);
    const segments: Segment[] = [];

    let start = 0;
    for (let i = 0; i < characters.length; i++) {
        const current = characters[i]!;
        const next = characters[i + 1];

        // A tab is alone in its segment: its width is a position rather than
        // a measurement, so it cannot be summed with the text beside it.
        const breakHere = current.mandatory
            || undefined === next
            || TAB === current.codePoint
            || TAB === next.codePoint
            || allowsBreakBetween(current.codePoint, next.codePoint);

        if (breakHere) {
            segments.push(makeSegment(runs, characters, start, i + 1, current.mandatory));
            start = i + 1;
        }
    }

    return segments;
}

interface FlatChar {
    readonly codePoint: number;
    readonly runIndex: number;
    readonly text: string;
    readonly mandatory: boolean;
    /**
     * A `w:softHyphen`: an OFFER of a break, and no ink until it is taken.
     *
     * Carried the way a mandatory break is — as a position rather than as a
     * character — because that is what it is. Keeping it in the text meant
     * every measurement had to remember to skip it, most fonts keep a hyphen
     * glyph at U+00AD ready to be drawn by accident, and a chop that cut in
     * front of one left it stranded alone on a line of its own.
     */
    readonly soft: boolean;
    /** Offset in the runs' concatenated text, in UTF-16 units. */
    readonly sourceIndex: number;
}

function flatten(runs: readonly StyledRun[]): FlatChar[] {
    const out: FlatChar[] = [];
    let sourceIndex = 0;

    runs.forEach((run, runIndex) => {
        for (const character of run.text) {
            const codePoint = character.codePointAt(0);
            if (undefined === codePoint) {
                continue;
            }
            out.push({
                codePoint,
                runIndex,
                text: character,
                // CRLF would otherwise end two lines where the author typed one.
                mandatory: isMandatoryBreak(codePoint),
                soft: SOFT_HYPHEN === codePoint,
                sourceIndex,
            });
            // Advance by UTF-16 units, not by one: iterating a string yields
            // whole code points, so an emoji or an astral character is one step
            // here and TWO units in every offset an editor works with.
            sourceIndex += character.length;
        }
    });

    // Collapse CRLF into a single break — by absorbing the line feed into the
    // carriage return rather than DROPPING it. Dropping ends one line where the
    // author typed one, but loses a unit of source: every offset after a Windows
    // line ending would then be one short, and the spans either side of the
    // break would no longer meet.
    //
    // The merged character stays mandatory, so its text is a position and is
    // never measured or drawn — only its LENGTH matters, and that is now 2.
    const collapsed: FlatChar[] = [];
    for (let index = 0; index < out.length; index++) {
        const character = out[index]!;
        const next = out[index + 1];

        if (0x000d === character.codePoint && undefined !== next && 0x000a === next.codePoint) {
            collapsed.push({ ...character, text: '\r\n' });
            index++;

            continue;
        }

        collapsed.push(character);
    }

    return collapsed;
}

function makeSegment(
    runs: readonly StyledRun[],
    characters: readonly FlatChar[],
    start: number,
    end: number,
    endsWithMandatoryBreak: boolean,
): Segment {
    // Group consecutive characters by run so each piece is measured with the
    // font it is actually drawn in.
    const pieces: LinePiece[] = [];
    let index = start;

    while (index < end) {
        const runIndex = characters[index]!.runIndex;
        const sourceStart = characters[index]!.sourceIndex;
        let text = '';

        while (index < end && characters[index]!.runIndex === runIndex) {
            const character = characters[index]!;
            // A mandatory break is a position, not ink: it must not be measured
            // and must not be rendered. Nor is a soft hyphen, until the line
            // ends at it and the piece is asked to show one.
            if (!character.mandatory && !character.soft) {
                text += character.text;
            }
            index++;
        }

        if ('' !== text) {
            const run = runs[runIndex]!;
            pieces.push({
                runIndex,
                text,
                // A picture has no glyphs: measuring its stand-in character
                // would give whatever the font keeps at .notdef.
                // A gloss wider than the text it sits over decides the advance,
                // and the text is centred under it when the renderer draws it.
                widthPx: undefined !== run.image
                    ? run.image.widthPx
                    : (undefined !== run.shape
                        // A shape may keep more of the line than it draws.
                        ? run.shape.advanceWidthPx ?? run.shape.widthPx
                        : Math.max(advanceOf(text, run), rubyWidthPx(run.ruby))),
                font: run.font,
                sizePx: run.sizePx,
                ...(undefined === run.colorHex ? {} : { colorHex: run.colorHex }),
                ...(undefined === run.highlightHex ? {} : { highlightHex: run.highlightHex }),
                ...(undefined === run.letterSpacingPx
                    ? {}
                    : { letterSpacingPx: run.letterSpacingPx }),
                ...(true === run.kerned ? { kerned: true } : {}),
                ...(undefined === run.footnoteId ? {} : { footnoteId: run.footnoteId }),
                ...(undefined === run.underline ? {} : { underline: run.underline }),
                ...(undefined === run.strike ? {} : { strike: run.strike }),
                ...(undefined === run.baselineShiftPx
                    ? {}
                    : { baselineShiftPx: run.baselineShiftPx }),
                ...(undefined === run.border
                    ? {}
                    : { border: run.border, borderRoomPx: borderRoomOf(run.border) }),
                ...(undefined === run.image ? {} : { image: run.image }),
                ...(undefined === run.shape ? {} : { shape: run.shape }),
                ...(undefined === run.ruby ? {} : { ruby: run.ruby }),
                sourceStart,
            });
        }
    }

    const first = characters[start]!;
    const last = characters[end - 1]!;

    const fullWidthPx = pieces.reduce((sum, piece) => sum + piece.widthPx, 0);
    const trailingSpaceWidthPx = measureTrailingSpaces(runs, characters, start, end);
    // ENDING a line at a soft hyphen costs the hyphen that is then drawn, where
    // carrying on past it costs nothing. The breaker already asks
    // those two questions separately — `content` is what it tests a break
    // against and `full` is what it charges for going on — so the offer needs
    // no machinery of its own.
    const offering = last.soft && pieces.length > 0;
    const hyphenPx = offering
        ? advanceOf('-', runs[pieces[pieces.length - 1]!.runIndex]!)
        : 0;

    if (offering) {
        pieces[pieces.length - 1] = { ...pieces[pieces.length - 1]!, hyphenAtEnd: true };
    }

    return {
        pieces,
        fullWidthPx,
        contentWidthPx: fullWidthPx - trailingSpaceWidthPx + hyphenPx,
        trailingSpaceWidthPx,
        endsWithMandatoryBreak,
        isTab: TAB === first.codePoint,
        sourceStart: first.sourceIndex,
        sourceEnd: last.sourceIndex + last.text.length,
    };
}

/**
 * Width of the spaces at the end of a segment.
 *
 * Measured separately rather than subtracted from a re-measurement, so that a
 * font whose space glyph kerns against its neighbour still adds up: the
 * difference between the segment with and without its spaces is exactly what
 * hangs into the margin.
 */
function measureTrailingSpaces(
    runs: readonly StyledRun[],
    characters: readonly FlatChar[],
    start: number,
    end: number,
): number {
    let width = 0;

    for (let i = end - 1; i >= start; i--) {
        const character = characters[i]!;
        if (character.mandatory) {
            continue;
        }
        if (!isSpace(character.codePoint) || !isBreakableSpace(character.codePoint)) {
            break;
        }

        const run = runs[character.runIndex]!;
        width += advanceOf(character.text, run);
    }

    return width;
}
