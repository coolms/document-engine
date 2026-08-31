import type { Alignment } from '../layout/alignment.js';
import type { BorderSide, BorderStyle, BoxBorders } from '../layout/borders.js';
import type { LineRule } from '../layout/baseline.js';
import type {
    Strike,
    VerticalAlign,
    TabAlign,
    TabLeader,
    Underline,
    UnderlineStyle,
} from '../layout/line-breaker.js';
import { eighthPointsToPx, pointsToPx } from '../ooxml/units.js';
import { XmlDocument, type XmlElement } from '../ooxml/xml.js';

/**
 * Where a paragraph's formatting actually comes from.
 *
 * Almost nothing in a Word document carries its own formatting. A heading is
 * `<w:pStyle w:val="Heading1"/>` and nothing else; its size lives in
 * `styles.xml`, on a style that is `basedOn` another style, which is based on a
 * third, and the bottom of that chain is `docDefaults`. Reading only the direct
 * properties gives every paragraph the same size and the document paginates
 * nothing like the original.
 *
 * ## Resolution order
 *
 * docDefaults, then the style chain from its ROOT downwards, then direct
 * formatting. Later wins. That order is what makes `basedOn` mean "inherit
 * from" rather than "override".
 *
 * ## Real files are broken in two specific ways
 *
 * A `basedOn` can name a style the file does not contain — the Word-authored
 * fixture here is `basedOn="Normal"` throughout and has no `Normal` style at
 * all. And a chain can contain a cycle. Neither may throw and neither may hang:
 * a document that Word opens has to open here.
 */

// One definition, in the layout layer that acts on it. Declaring the union
// here as well would let the two drift, and a rule the reader can express
// but the baseline cannot is a silently mispositioned line.
export type { LineRule } from '../layout/baseline.js';

export interface RunProperties {
    readonly family?: string;
    /** `w:sz` — HALF-points, so 20 is 10pt. */
    readonly sizeHalfPoints?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    /**
     * `w:color/@w:val` as `#RRGGBB`.
     *
     * `auto` resolves to black rather than staying absent. It is a VALUE, not
     * the absence of one: a run carrying it overrides the colour its style
     * supplies, so leaving it absent would let an inherited colour show through
     * where the document asked for automatic.
     *
     * Verified against LibreOffice: a run with a character style setting green
     * and its own `w:val="auto"` prints BLACK, while the same run without the
     * `auto` prints green.
     */
    readonly colorHex?: string;
    /**
     * The colour painted BEHIND the run — `w:highlight` or `w:shd/@w:fill`.
     *
     * One field for two elements because they produce the same mark: a marker
     * pen and a shaded run both fill the line box behind the text, and a
     * renderer given them separately would draw one over the other.
     */
    readonly highlightHex?: string;
    /** `w:kern` — see {@link readKerning}. Absent means not kerned. */
    readonly kerned?: boolean;
    /** `w:u` — a rule under the text. */
    readonly underline?: Underline;
    /** `w:strike` / `w:dstrike` — a rule through it. */
    readonly strike?: Strike;
    /** `w:vertAlign` — off the baseline, and smaller. */
    readonly vertAlign?: VerticalAlign;
    /**
     * `w:position` — off the baseline by hand, in half-points, and NOT smaller.
     *
     * Half-points of the LINE, not of the point size: measured against
     * LibreOffice, twelve units raised a run 6.90pt in Liberation Serif, 6.75
     * in Liberation Mono and 7.00 in DejaVu Sans — the three fonts' own
     * line-height ratios, and unchanged by doubling the size or pinning the
     * line to an exact 24pt.
     */
    readonly positionHalfPoints?: number;
    /**
     * `w:bdr` — one border, drawn round all four sides of the run.
     *
     * A single side rather than a {@link BoxBorders}: `w:bdr` is ONE element
     * with one width, one colour and one space, and Word draws it round the
     * whole run. Measured against LibreOffice, the box stands `w:space` plus
     * half the rule outside the text on every side, and the run reserves the
     * space plus the WHOLE rule — a boxed run in an 11.5pt line made the line
     * 21.5 for a three-point rule two points clear.
     */
    readonly border?: BorderSide;
    /** `w:caps` — draw the run in capitals whatever the file stores. */
    readonly caps?: boolean;
    /** `w:smallCaps` — draw the LOWER case as smaller capitals. */
    readonly smallCaps?: boolean;
    /** `w:spacing` on a RUN — tracking, in twentieths of a point. */
    readonly letterSpacingTwentieths?: number;
    /**
     * `w:vanish` — text the document says to HIDE.
     *
     * Measured against LibreOffice: a hidden run between two visible
     * ones is not drawn and takes no room either — the words either side sit
     * 19.45 apart, which is the first one's own advance. So it is dropped
     * rather than made invisible, and a reader that ignores it prints what the
     * author hid.
     */
    readonly hidden?: boolean;
}

export interface ParagraphProperties {
    readonly spaceBeforeTwips?: number;
    readonly spaceAfterTwips?: number;
    readonly lineTwips?: number;
    readonly lineRule?: LineRule;
    readonly indentLeftTwips?: number;
    readonly indentRightTwips?: number;
    /** Positive from `w:firstLine`, negative from `w:hanging`. */
    readonly indentFirstLineTwips?: number;
    readonly pageBreakBefore?: boolean;
    readonly keepLines?: boolean;
    /** `w:keepNext` — stay on the same page as the block that follows. */
    readonly keepNext?: boolean;
    /** `w:pBdr` — the paragraph's own border. */
    readonly borders?: BoxBorders;
    readonly widowControl?: boolean;
    /**
     * `w:contextualSpacing` — spend no space against a NEIGHBOUR of the same
     * style.
     *
     * Word's list styles set it as a matter of course, which is why a list is
     * tight and the paragraphs either side of it are not. Measured against
     * LibreOffice: a pair that prints 21.50 apart without it prints
     * 11.50 with it, and the space is dropped against a plain neighbour of the
     * same style too — the flag belongs to the paragraph that gives up its own
     * space, not to the pair.
     */
    readonly contextualSpacing?: boolean;
    /**
     * `w:suppressLineNumbers` — take this paragraph out of `w:lnNumType`.
     *
     * Measured against LibreOffice: the paragraph before it printed
     * line 10, this one printed no number at all, and the next took 11 — so
     * the line is not counted rather than counted and left blank.
     */
    readonly suppressLineNumbers?: boolean;
    /**
     * `w:snapToGrid` — whether this paragraph sits on the section's
     * `w:docGrid` at all.
     *
     * Absent is ON, and Word writes the OFF form freely, so a gridded document
     * is likely to hold some. Measured against LibreOffice: under an
     * 18pt grid a paragraph that says no steps 11.50, the font's own line,
     * where its neighbour on the grid steps 18.00 — and the grid does not
     * re-snap afterwards, the same way an `exact` line rule escapes it.
     */
    readonly snapToGrid?: boolean;
    /** `w:tabs` — explicit stops in twips, from the paragraph's left indent. */
    readonly tabStops?: readonly {
        positionTwips: number;
        align: TabAlign;
        leader?: TabLeader;
    }[];
    /** `w:jc`, normalised — see {@link readAlignment}. */
    readonly alignment?: Alignment;
}

/** What a style says about list membership, if anything. */
export interface StyleNumbering {
    readonly numId: string;
    readonly ilvl: number;
}

/** One `w:tblStylePr`: what a table style lends to a part of the table. */
export interface ConditionalFormat {
    readonly paragraph: ParagraphProperties;
    readonly run: RunProperties;
    /** Raw, because the cell properties are read by the document reader. */
    readonly tcPr: XmlElement | null;
}

function readConditions(style: XmlElement): Map<string, ConditionalFormat> {
    const out = new Map<string, ConditionalFormat>();
    for (const conditional of style.elements('w:tblStylePr')) {
        const type = conditional.attribute('w:type');
        if (null === type) {
            continue;
        }

        out.set(type, {
            paragraph: readParagraphProperties(conditional.element('w:pPr') ?? null),
            run: readRunProperties(conditional.element('w:rPr') ?? null),
            tcPr: conditional.element('w:tcPr') ?? null,
        });
    }

    return out;
}

interface StyleEntry {
    readonly id: string;
    /** `w:name` — what the style is CALLED, which is how a built-in is known. */
    readonly name: string | null;
    readonly basedOn: string | null;
    readonly paragraph: ParagraphProperties;
    readonly run: RunProperties;
    readonly numbering: StyleNumbering | null;
}

export class StyleSheet {
    private readonly paragraphStyles = new Map<string, StyleEntry>();
    private readonly characterStyles = new Map<string, StyleEntry>();
    /**
     * Table styles, kept as the RAW `w:tblPr` of each.
     *
     * Unparsed because a table's properties are read by the document reader,
     * where the border and margin readers live — and because the merge is
     * per SIDE across the chain rather than per element, which needs the
     * elements themselves.
     */
    private readonly tableStyles = new Map<string, {
        basedOn: string | null;
        tblPr: XmlElement | null;
        paragraph: ParagraphProperties;
        run: RunProperties;
        conditions: Map<string, ConditionalFormat>;
    }>();

    private constructor(
        private readonly defaultParagraph: ParagraphProperties,
        private readonly defaultRun: RunProperties,
        private defaultParagraphStyleId: string | null,
    ) {}

    /** An absent `styles.xml` is legal; the document then gets bare defaults. */
    static empty(): StyleSheet {
        return new StyleSheet({}, {}, null);
    }

    static parse(source: string): StyleSheet {
        const root = XmlDocument.parse(source).root;

        const defaults = root.element('w:docDefaults');
        const sheet = new StyleSheet(
            readParagraphProperties(defaults?.element('w:pPrDefault')?.element('w:pPr') ?? null),
            withStatedDefaults(
                readRunProperties(defaults?.element('w:rPrDefault')?.element('w:rPr') ?? null)),
            null,
        );

        for (const style of root.elements('w:style')) {
            const id = style.attribute('w:styleId');
            if (null === id) {
                continue;
            }

            const entry: StyleEntry = {
                id,
                name: style.element('w:name')?.attribute('w:val') ?? null,
                basedOn: style.element('w:basedOn')?.attribute('w:val') ?? null,
                paragraph: readParagraphProperties(style.element('w:pPr')),
                run: readRunProperties(style.element('w:rPr')),
                numbering: readStyleNumbering(style.element('w:pPr')),
            };

            const type = style.attribute('w:type');
            if ('table' === type) {
                sheet.tableStyles.set(id, {
                    basedOn: entry.basedOn,
                    tblPr: style.element('w:tblPr') ?? null,
                    paragraph: entry.paragraph,
                    run: entry.run,
                    conditions: readConditions(style),
                });
            } else if ('character' === type) {
                sheet.characterStyles.set(id, entry);
            } else if ('paragraph' === type || null === type) {
                sheet.paragraphStyles.set(id, entry);
                if ('1' === style.attribute('w:default') && null === sheet.defaultParagraphStyleId) {
                    sheet.defaultParagraphStyleId = id;
                }
            }
        }

        return sheet;
    }

    /**
     * The list a style puts its paragraphs in, if any.
     *
     * Word's own "List Paragraph" works this way: the paragraph names only a
     * style, and the style says which list it belongs to. A reader that looks
     * only at `w:numPr` on the paragraph sees no list at all.
     *
     * Resolved down the `basedOn` chain, so a style inherits its list from
     * the one it is based on.
     */
    numbering(styleId: string | null): StyleNumbering | null {
        let out: StyleNumbering | null = null;
        for (const entry of this.chain(styleId ?? this.defaultParagraphStyleId, this.paragraphStyles)) {
            out = entry.numbering ?? out;
        }

        return out;
    }

    /**
     * The `w:tblPr` of every style this table style is built from, ROOT FIRST.
     *
     * Root first because each level answers for the sides it names and the ones
     * after it override — see `readBordersFrom` in the document reader, which
     * is given this chain with the table's own properties on the end.
     */
    tableStyleProperties(styleId: string | null): XmlElement[] {
        const out: XmlElement[] = [];
        const seen = new Set<string>();

        let current = styleId;
        while (null !== current && !seen.has(current)) {
            seen.add(current);
            const entry = this.tableStyles.get(current);
            if (undefined === entry) {
                break;
            }
            if (null !== entry.tblPr) {
                out.push(entry.tblPr);
            }
            current = entry.basedOn;
        }

        return out.reverse();
    }

    /**
     * The paragraph and run properties a table style lends to the text INSIDE
     * it, merged root-first down its `basedOn` chain.
     *
     * Measured against LibreOffice: a style whose own `w:rPr` says
     * fourteen points stepped its rows 16.60 where the same table without one
     * stepped 12.05, and a style whose `w:pPr` says centred put its text at
     * 183.05 rather than 77.25. So a table style carries more than the table's
     * own rules — it dresses the text as well.
     *
     * They sit UNDER the paragraph's own style, which is where the format puts
     * them: docDefaults, then the table style, then the paragraph's.
     */
    tableStyleCascade(styleId: string | null): {
        paragraph: ParagraphProperties;
        run: RunProperties;
    } {
        let paragraph: ParagraphProperties = {};
        let run: RunProperties = {};
        const seen = new Set<string>();
        const chain: string[] = [];

        let current = styleId;
        while (null !== current && !seen.has(current)) {
            seen.add(current);
            if (undefined === this.tableStyles.get(current)) {
                break;
            }
            chain.push(current);
            current = this.tableStyles.get(current)?.basedOn ?? null;
        }

        for (const id of chain.reverse()) {
            const entry = this.tableStyles.get(id)!;
            paragraph = { ...paragraph, ...entry.paragraph };
            run = { ...run, ...entry.run };
        }

        return { paragraph, run };
    }

    /**
     * The conditional formats a table style declares, merged down its chain.
     *
     * Keyed by `w:tblStylePr/@w:type` — `firstRow`, `lastRow`, `firstCol`,
     * `lastCol`, `band1Horz`, `band2Horz` and the rest. Which CELLS each one
     * reaches, and which of them wins where they meet, is the document
     * reader's business; this only says what the style declares.
     */
    tableStyleConditions(styleId: string | null): Map<string, ConditionalFormat> {
        const merged = new Map<string, ConditionalFormat>();
        const seen = new Set<string>();
        const chain: string[] = [];

        let current = styleId;
        while (null !== current && !seen.has(current)) {
            seen.add(current);
            if (undefined === this.tableStyles.get(current)) {
                break;
            }
            chain.push(current);
            current = this.tableStyles.get(current)?.basedOn ?? null;
        }

        for (const id of chain.reverse()) {
            for (const [type, format] of this.tableStyles.get(id)!.conditions) {
                const under = merged.get(type);
                merged.set(type, undefined === under ? format : {
                    paragraph: { ...under.paragraph, ...format.paragraph },
                    run: { ...under.run, ...format.run },
                    tcPr: format.tcPr ?? under.tcPr,
                });
            }
        }

        return merged;
    }

    /** True when the document can name this paragraph style and get something. */
    hasParagraphStyle(id: string): boolean {
        return this.paragraphStyles.has(id);
    }

    /**
     * Whether this style IS one of the built-in headings, by the name the
     * format gives them — `heading 1` through `heading 9`.
     *
     * By the NAME and not by the id, which is measured rather than assumed: a
     * style called `Zebra` and named `heading 6` was spaced like a heading, and
     * one called `Heading7` and named `Custom Thing` was not.
     */
    isBuiltInHeading(styleId: string | null): boolean {
        const name = null === styleId ? null : this.paragraphStyles.get(styleId)?.name;

        return undefined !== name && null !== name && /^heading [1-9]$/i.test(name.trim());
    }

    paragraphProperties(
        styleId: string | null,
        under: ParagraphProperties = {},
    ): ParagraphProperties {
        let out: ParagraphProperties = { ...this.defaultParagraph, ...under };
        for (const entry of this.chain(styleId ?? this.defaultParagraphStyleId, this.paragraphStyles)) {
            out = { ...out, ...entry.paragraph };
        }

        return out;
    }

    /**
     * A run's properties: document defaults, then the PARAGRAPH style's run
     * properties, then any character style. A heading's size lives on the
     * paragraph style's `w:rPr`, so skipping that step loses it.
     */
    runProperties(
        paragraphStyleId: string | null,
        characterStyleId: string | null,
        under: RunProperties = {},
    ): RunProperties {
        let out: RunProperties = { ...this.defaultRun, ...under };
        for (const entry of this.chain(paragraphStyleId ?? this.defaultParagraphStyleId, this.paragraphStyles)) {
            out = { ...out, ...entry.run };
        }
        for (const entry of this.chain(characterStyleId, this.characterStyles)) {
            out = { ...out, ...entry.run };
        }

        return out;
    }

    /**
     * The `basedOn` ancestry, ROOT FIRST so applying in order lets a descendant
     * override its base.
     *
     * A name that resolves to nothing simply ends the walk — that is what Word
     * does, and the alternative is refusing to open a document it opens fine. A
     * cycle ends it too, which is what the `seen` set is for: without it, a
     * self-referential style is an infinite loop rather than a bad style.
     */
    private chain(styleId: string | null, styles: Map<string, StyleEntry>): StyleEntry[] {
        const out: StyleEntry[] = [];
        const seen = new Set<string>();

        let current = styleId;
        while (null !== current && !seen.has(current)) {
            seen.add(current);
            const entry = styles.get(current);
            if (undefined === entry) {
                break;
            }
            out.push(entry);
            current = entry.basedOn;
        }

        return out.reverse();
    }
}

/**
 * `w:tabs`, in ascending order.
 *
 * A `w:val="clear"` stop REMOVES an inherited one rather than adding a
 * column, so it is dropped here. A stop's OWN alignment (`w:val="center"`,
 * `"right"`, `"decimal"`) is still read as left, which puts the text in the
 * right column and starts it on the wrong side of that column — separate from
 * the PARAGRAPH alignment that {@link readAlignment} handles.
 */
function readTabStops(
    pPr: XmlElement | null,
): { positionTwips: number; align: TabAlign }[] | null {
    const tabs = pPr?.element('w:tabs') ?? null;
    if (null === tabs) {
        return null;
    }

    const stops: { positionTwips: number; align: TabAlign }[] = [];
    for (const stop of tabs.elements('w:tab')) {
        const value = stop.attribute('w:val') ?? 'left';
        // `clear` REMOVES an inherited stop rather than adding a column, and
        // `bar` draws a vertical rule without being a stop at all — a tab never
        // advances to one.
        if ('clear' === value || 'bar' === value) {
            continue;
        }

        const position = Number(stop.attribute('w:pos') ?? '');
        if (Number.isFinite(position)) {
            const leader = LEADERS[stop.attribute('w:leader') ?? 'none'];
            stops.push({
                positionTwips: position,
                align: TAB_ALIGNS[value] ?? 'left',
                ...(undefined === leader ? {} : { leader }),
            });
        }
    }

    return 0 === stops.length ? null : stops.sort((a, b) => a.positionTwips - b.positionTwips);
}

/**
 * `w:tab/@w:val`, folded onto what the breaker can do.
 *
 * `num` is a list-numbering stop that behaves as a left one, and anything
 * unrecognised is left too — a stop in the right column and aligned the usual
 * way is far closer to the document than no stop at all.
 */
/**
 * `w:tab/@w:leader`.
 *
 * `none` is the default and the absence of a leader, so it is simply missing
 * from this map — as is anything unrecognised, which draws a plain tab rather
 * than a row of the wrong character.
 */
const LEADERS: Record<string, TabLeader | undefined> = {
    dot: 'dot',
    hyphen: 'hyphen',
    underscore: 'underscore',
    middleDot: 'middleDot',
    heavy: 'heavy',
};

const TAB_ALIGNS: Record<string, TabAlign> = {
    left: 'left',
    start: 'left',
    num: 'left',
    center: 'center',
    right: 'right',
    end: 'right',
    decimal: 'decimal',
};

/**
 * `w:jc`, normalised.
 *
 * Word writes `left`/`right`; the strict schema writes `start`/`end`; both mean
 * the same thing here because this engine has no bidirectional text, so there is
 * no case where the start of a line is not its left. `distribute` stretches the
 * last line as well as the others and is read as `justify`, which does not —
 * the difference shows on one line per paragraph.
 *
 * Anything unrecognised returns null rather than a guess: answering `left`
 * would silently overrule an alignment a parent style got right.
 */
function readAlignment(pPr: XmlElement | null): Alignment | null {
    switch (pPr?.element('w:jc')?.attribute('w:val') ?? null) {
        case 'left':
        case 'start':
            return 'left';
        case 'center':
            return 'center';
        case 'right':
        case 'end':
            return 'right';
        case 'both':
        case 'distribute':
            return 'justify';
        default:
            return null;
    }
}

/** The list a style's own `w:numPr` puts its paragraphs in, if any. */
function readStyleNumbering(pPr: XmlElement | null | undefined): StyleNumbering | null {
    const numPr = pPr?.element('w:numPr') ?? null;
    const numId = numPr?.element('w:numId')?.attribute('w:val') ?? null;
    if (null === numId) {
        return null;
    }

    const ilvl = Number(numPr?.element('w:ilvl')?.attribute('w:val') ?? '0');

    return { numId, ilvl: Number.isFinite(ilvl) ? ilvl : 0 };
}

/**
 * Read the properties written directly on a `w:pPr`.
 *
 * Exported because the document reader needs the SAME parse for a paragraph's
 * own formatting as for a style's. Two parsers would drift, and the direction
 * of the drift would be a precedence bug rather than a parse error.
 */
export function readParagraphProperties(pPr: XmlElement | null): ParagraphProperties {
    if (null === pPr) {
        return {};
    }

    const spacing = pPr.element('w:spacing');
    const indent = pPr.element('w:ind');
    const lineRule = spacing?.attribute('w:lineRule') ?? null;

    // `w:hanging` is a first-line indent pointing the other way, and the two
    // are mutually exclusive. Reading only firstLine turns every hanging indent
    // — which is what a bulleted list uses — into no indent at all.
    const firstLine = number(indent, 'w:firstLine');
    const hanging = number(indent, 'w:hanging');

    return {
        ...pick('spaceBeforeTwips', number(spacing, 'w:before')),
        ...pick('spaceAfterTwips', number(spacing, 'w:after')),
        ...pick('lineTwips', number(spacing, 'w:line')),
        ...pick('lineRule', isLineRule(lineRule) ? lineRule : null),
        // w:start and w:end are the current spelling of w:left and w:right.
        ...pick('indentLeftTwips', number(indent, 'w:left') ?? number(indent, 'w:start')),
        ...pick('indentRightTwips', number(indent, 'w:right') ?? number(indent, 'w:end')),
        ...pick('indentFirstLineTwips', null !== hanging ? -hanging : firstLine),
        ...pick('pageBreakBefore', toggle(pPr, 'w:pageBreakBefore')),
        ...pick('keepLines', toggle(pPr, 'w:keepLines')),
        ...pick('keepNext', toggle(pPr, 'w:keepNext')),
        ...pick('widowControl', toggle(pPr, 'w:widowControl')),
        ...pick('contextualSpacing', toggle(pPr, 'w:contextualSpacing')),
        ...pick('suppressLineNumbers', toggle(pPr, 'w:suppressLineNumbers')),
        ...pick('snapToGrid', toggle(pPr, 'w:snapToGrid')),
        ...pick('tabStops', readTabStops(pPr)),
        ...pick('alignment', readAlignment(pPr)),
    };
}

/**
 * What a `styles.xml` that states no default FAMILY means: Times New Roman.
 *
 * ## Which is not the same as a package with no styles part at all
 *
 * The two look alike and print differently, measured with the font's own name
 * read out of the printed PDF rather than inferred from its metrics:
 *
 *   * a `.docx` holding `document.xml` and nothing else embedded
 *     `Carlito-Regular` — Calibri, which is what the font manifest's `defaults`
 *     answers and what this engine already did;
 *   * a `.docx` whose `styles.xml` carries an EMPTY `w:rPrDefault` embedded
 *     `LiberationSerif` — Times New Roman, and nothing else.
 *
 * The second is not a corner: it is what an UPLOADED template that states
 * nothing looks like, and the fill path copies its `styles.xml` verbatim.
 * Answering Calibri there made such a document paginate one way in the preview
 * and another way in the renderer that prints it — 1.2207 em a line against
 * 1.15, which is a line every eighteen.
 *
 * It is NOT what this platform writes, though the file that gave it away looked
 * like proof: `coolms2-adr-029.docx` came from the `docx` npm package, which is
 * not a dependency here. Everything this codebase composes goes through PHPWord,
 * which always states its run defaults.
 *
 * Named as the FACE rather than as `Times New Roman`, though the manifest maps
 * one to the other, because the document never asked for Times New Roman: this
 * engine chose it. Routed through the substitution path it would collect a
 * "not available, using the metric-compatible…" notice on every document that
 * omits a font — and a diagnostic that fires on everything is worse than none.
 * A document that names the font itself still gets that notice, which is the
 * case it is for.
 */
const NO_STATED_FAMILY = 'Liberation Serif';

/**
 * And its size: ten points, which is the other half of the same fallback.
 *
 * Measured in the same way. A styles part with an empty
 * `w:rPrDefault`, paragraphs stating no `w:rPr` at all, and a 10pt paragraph
 * beside them as a ruler in the same print: both stepped 11.55, so a run that
 * says nothing is set at ten points.
 *
 * The bare package answers differently again and needs nothing here — its
 * unstated runs stepped 13.45, which is eleven points of Carlito's 1.2207 em,
 * the manifest's `defaults` exactly. Both halves of that fallback are right
 * where they are; both halves of this one are wrong anywhere else.
 */
const NO_STATED_SIZE_HALF_POINTS = 20;

function withStatedDefaults(properties: RunProperties): RunProperties {
    return {
        ...properties,
        ...(undefined === properties.family ? { family: NO_STATED_FAMILY } : {}),
        ...(undefined === properties.sizeHalfPoints
            ? { sizeHalfPoints: NO_STATED_SIZE_HALF_POINTS }
            : {}),
    };
}

export function readRunProperties(rPr: XmlElement | null): RunProperties {
    if (null === rPr) {
        return {};
    }

    const fonts = rPr.element('w:rFonts');

    return {
        // w:ascii is the Latin face. w:cs and w:eastAsia cover scripts this
        // engine does not shape yet, and taking them would pick the wrong font
        // for ordinary text.
        ...pick('family', fonts?.attribute('w:ascii') ?? fonts?.attribute('w:hAnsi') ?? null),
        ...pick('sizeHalfPoints', number(rPr.element('w:sz'), 'w:val')),
        ...pick('bold', toggle(rPr, 'w:b')),
        ...pick('italic', toggle(rPr, 'w:i')),
        ...pick('colorHex', readColor(rPr.element('w:color'))),
        // A marker pen wins over a shaded background: Word draws the
        // highlight over the shading, so the shading is never seen.
        ...pick('highlightHex', readHighlight(rPr.element('w:highlight'))
            ?? readShadingFill(rPr.element('w:shd'))),
        ...pick('underline', readUnderline(rPr.element('w:u'))),
        ...pick<'vertAlign', VerticalAlign>('vertAlign',
            readVerticalAlign(rPr.element('w:vertAlign'))),
        ...pick('positionHalfPoints', number(rPr.element('w:position'), 'w:val')),
        ...pick<'border', BorderSide>('border', readBorderSide(rPr.element('w:bdr'))),
        ...pick('caps', toggle(rPr, 'w:caps')),
        ...pick('smallCaps', toggle(rPr, 'w:smallCaps')),
        ...pick('hidden', toggle(rPr, 'w:vanish')),
        ...pick('letterSpacingTwentieths', number(rPr.element('w:spacing'), 'w:val')),
        ...pick('kerned', readKerning(rPr.element('w:kern'))),
        // A DOUBLE strike wins: a run carrying both is struck twice in
        // Word, not once.
        ...pick<'strike', Strike>('strike',
            true === toggle(rPr, 'w:dstrike')
                ? 'double'
                : (true === toggle(rPr, 'w:strike') ? 'single' : null)),
    };
}

/**
 * `w:kern`: whether this run's pairs are kerned at all.
 *
 * The element states the SMALLEST size kerning applies to, in half-points, and
 * zero is how the format spells "off". LibreOffice reads only whether it is
 * there: a run saying `w:kern="40"` — kern nothing under 20pt — printed its
 * 10pt text kerned all the same, identically to one saying `16`.
 * The threshold is therefore not read here either, because honouring it would
 * put us at odds with the page for exactly the documents that state one.
 *
 * Absent means NOT kerned, which is what makes this worth reading at all:
 * nothing this platform generates says `w:kern`, and every one of those runs
 * was being measured kerned.
 */
function readKerning(kern: XmlElement | null): boolean | null {
    if (null === kern) {
        return null;
    }

    const threshold = number(kern, 'w:val');

    return null === threshold || threshold > 0;
}

const BORDER_STYLES: Record<string, BorderStyle> = {
    dashed: 'dashed',
    dashSmallGap: 'dashed',
    dotted: 'dotted',
    dotDash: 'dashed',
    dotDotDash: 'dashed',
    double: 'double',
    triple: 'double',
};

/** Word's default when a border declares a style but no width: half a point. */
const DEFAULT_BORDER_EIGHTHS = 4;

export function readBorderSide(element: XmlElement | null): BorderSide | null {
    const value = element?.attribute('w:val') ?? null;
    // `nil` and `none` are how a border is switched OFF, and are not the same
    // as the element being absent.
    if (null === value || 'nil' === value || 'none' === value) {
        return null;
    }

    const eighths = number(element, 'w:sz');
    const colour = element?.attribute('w:color') ?? null;
    // `w:space` counts POINTS, where the width beside it counts eighths.
    const space = number(element, 'w:space');

    return {
        ...(null === space || 0 === space ? {} : { spacePx: pointsToPx(space) }),
        widthPx: eighthPointsToPx(
            null !== eighths && eighths > 0 ? eighths : DEFAULT_BORDER_EIGHTHS,
        ),
        // `auto` means "let the consumer decide", and every consumer decides
        // black — including Word.
        colorHex: null === colour || 'auto' === colour ? '#000000' : `#${colour}`,
        style: BORDER_STYLES[value] ?? 'solid',
    };
}

/**
 * `w:vertAlign` — whether a run sits off the baseline.
 *
 * `baseline` is the explicit "on the line", which a run uses to CANCEL what its
 * style asked for, so it has to be a value rather than an absence.
 */
function readVerticalAlign(element: XmlElement | null): VerticalAlign | null {
    const value = element?.attribute('w:val') ?? null;

    return 'superscript' === value || 'subscript' === value || 'baseline' === value
        ? value
        : null;
}

/**
 * `w:u` — the rule under a run.
 *
 * `none` is how underlining is switched OFF and yields null, which lets a run
 * cancel what its style asked for. An unrecognised value still draws a SINGLE
 * rule: the presence of the underline is what a reader sees, and dropping it
 * would lose the emphasis entirely.
 */
function readUnderline(element: XmlElement | null): Underline | null {
    const value = element?.attribute('w:val') ?? null;
    if (null === value || 'none' === value) {
        return null;
    }

    const colour = element?.attribute('w:color') ?? null;

    return {
        style: UNDERLINES[value] ?? 'single',
        ...(null === colour || 'auto' === colour ? {} : { colorHex: `#${colour}` }),
    };
}

const UNDERLINES: Record<string, UnderlineStyle | undefined> = {
    single: 'single',
    words: 'single',
    thick: 'single',
    double: 'double',
    dotted: 'dotted',
    dottedHeavy: 'dotted',
    dash: 'dashed',
    dashedHeavy: 'dashed',
    dashLong: 'dashed',
    dotDash: 'dashed',
    dotDotDash: 'dashed',
    wave: 'single',
    wavyHeavy: 'single',
    wavyDouble: 'double',
};

/**
 * `w:highlight` — a marker pen, named rather than given as a colour.
 *
 * The seventeen names are fixed by the schema and are the only values allowed,
 * so an unrecognised one is a producer's mistake and is ignored rather than
 * guessed at. `none` is how a highlight is switched OFF and yields null.
 *
 * The hexes were checked against LibreOffice's own output: `yellow` prints as
 * `1 1 0 rg` and `green` as `0 1 0 rg`.
 */
function readHighlight(element: XmlElement | null): string | null {
    return HIGHLIGHTS[element?.attribute('w:val') ?? ''] ?? null;
}

const HIGHLIGHTS: Record<string, string> = {
    black: '#000000',
    blue: '#0000FF',
    cyan: '#00FFFF',
    darkBlue: '#000080',
    darkCyan: '#008080',
    darkGray: '#808080',
    darkGreen: '#008000',
    darkMagenta: '#800080',
    darkRed: '#800000',
    darkYellow: '#808000',
    green: '#00FF00',
    lightGray: '#C0C0C0',
    magenta: '#FF00FF',
    red: '#FF0000',
    white: '#FFFFFF',
    yellow: '#FFFF00',
};

/**
 * A run's `w:shd/@w:fill` — the same rule the TABLE reader uses for cells.
 *
 * `w:color` on the same element is a hatch pattern's foreground, which this
 * engine does not draw; taking it would fill the run with the wrong colour.
 */
function readShadingFill(element: XmlElement | null): string | null {
    const fill = element?.attribute('w:fill') ?? null;

    return null === fill || 'auto' === fill ? null : `#${fill}`;
}

/**
 * `w:color`, as `#RRGGBB`.
 *
 * `auto` means "a colour that reads against the background", and on the white
 * paper this engine draws that is black — which is what LibreOffice prints for
 * it. True contrast resolution against a shaded cell is NOT done, so automatic
 * text on a dark fill comes out black rather than white.
 *
 * A `w:themeColor` is not resolved: the theme lives in `theme1.xml`, which this
 * reader is not given, and inventing a colour for it would be worse than
 * inheriting one. Such a run keeps whatever its style supplies.
 */
function readColor(element: XmlElement | null): string | null {
    const value = element?.attribute('w:val') ?? null;
    if (null === value) {
        return null;
    }

    return 'auto' === value ? '#000000' : `#${value}`;
}

/**
 * Include a key only when there is a value.
 *
 * A key present but undefined would override an inherited value during the
 * merge — it has to be ABSENT for inheritance to work at all.
 */
function pick<K extends string, V>(key: K, value: V | null): { [P in K]?: V } {
    return (null === value ? {} : { [key]: value }) as { [P in K]?: V };
}

function number(element: XmlElement | null | undefined, attribute: string): number | null {
    const raw = element?.attribute(attribute) ?? null;
    if (null === raw) {
        return null;
    }
    const value = Number(raw);

    return Number.isFinite(value) ? value : null;
}

/**
 * A WordprocessingML toggle: present means on, `w:val="0"` (or false, or off)
 * means explicitly OFF, and absent means inherit.
 *
 * The off case is not decoration — it is how a paragraph escapes a style that
 * turns something on, and reading presence alone makes it impossible to turn
 * anything off.
 */
function toggle(parent: XmlElement, name: string): boolean | null {
    const element = parent.element(name);
    if (null === element) {
        return null;
    }

    const value = element.attribute('w:val');
    if (null === value) {
        return true;
    }

    return !('0' === value || 'false' === value || 'off' === value);
}

function isLineRule(value: string | null): value is LineRule {
    return 'auto' === value || 'exact' === value || 'atLeast' === value;
}
