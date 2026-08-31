import { formatNumeral } from '../text/numerals.js';
import { XmlDocument, type XmlElement } from '../ooxml/xml.js';

/**
 * Lists: `numbering.xml`, and the counters that turn it into "1.", "a)" or "●".
 *
 * ## Two levels of indirection, and the COUNT lives at the bottom of them
 *
 * A paragraph names a `w:numId`. That points at an ABSTRACT numbering, and
 * several numIds can share one — but they do not then count separately.
 * Measured against LibreOffice: three numIds on one abstract, used in
 * turn, counted straight through it — `(C) (D)`, then `5. 6.`, then `7. 8.`,
 * then `9. 10.` — so the running number belongs to the ABSTRACT and a change of
 * numId does not restart it.
 *
 * What restarts it is an OVERRIDE, and that is why Word writes one every time
 * you ask a list to begin again. `w:startOverride` sets the count, ONCE, at the
 * first paragraph that uses its numId: the same numId's later items carried on
 * from wherever the document had reached — `11. 12.` where a second reset would
 * have said `7. 8.` again.
 *
 * A `w:lvl` inside the override is a different thing: it replaces the level's
 * definition — its format, its pattern, its indent — for the instance that
 * states it. Its `w:start` is an initial value like any level's, not a reset:
 * the same instance printed `(C)` from a `w:start` of 3 with the counter fresh,
 * and `(J)` from the very same numbering when the counter had already run to 9.
 *
 * ## The marker is not text
 *
 * "1." is not part of the paragraph's content: it is drawn in the hanging indent
 * beside the first line, and the text wraps to the LEFT indent underneath it.
 * That is why a list's indent comes from the numbering level rather than from
 * the paragraph, and why the marker never affects where a line breaks.
 */

export type NumberFormat =
    | 'bullet'
    | 'decimal'
    | 'lowerLetter'
    | 'upperLetter'
    | 'lowerRoman'
    | 'upperRoman'
    | 'none';

export interface NumberingLevel {
    readonly format: NumberFormat;
    /** `w:lvlText`: a bullet glyph, or a pattern like "%1." or "%1.%2". */
    readonly text: string;
    readonly start: number;
    readonly indentLeftTwips: number | null;
    readonly indentHangingTwips: number | null;
    /** The font the bullet GLYPH is drawn in — Symbol and Wingdings are common. */
    readonly fontFamily: string | null;
    /** `w:suff`: what separates the marker from the text. Word's default is a tab. */
    readonly suffix: 'tab' | 'space' | 'nothing';
    /**
     * `w:lvlJc`: how the marker sits against the point the hanging indent
     * gives it — NOT how it fills the space between that point and the text.
     *
     * Measured: all three justifications place the marker around the
     * SAME anchor, `indentLeft - hanging`. `left` starts there, `right` ends
     * there and grows away from the text, `center` straddles it.
     */
    readonly justification: 'left' | 'right' | 'center';
    /**
     * `w:lvlRestart`. Zero means this level NEVER restarts — its count runs
     * on through the whole document however often the levels above it change.
     */
    readonly restart: number | null;
}

interface AbstractNumbering {
    readonly levels: Map<number, NumberingLevel>;
}

/** One `w:num`: which definition it uses, and what it changes about it. */
interface NumberingInstance {
    readonly abstractId: string;
    /** ilvl to this instance's `w:lvlOverride` for that level. */
    readonly overrides: Map<number, LevelOverride>;
}

interface LevelOverride {
    /** `w:startOverride` — the count is SET to this, once. See {@link Numbering}. */
    readonly startAt: number | null;
    /** A `w:lvl` inside the override, replacing the definition for this instance. */
    readonly level: NumberingLevel | null;
}

export class Numbering {
    private constructor(
        private readonly abstracts: Map<string, AbstractNumbering>,
        private readonly instances: Map<string, NumberingInstance>,
    ) {}

    /** A document without `numbering.xml` — legal, and it simply has no lists. */
    static empty(): Numbering {
        return new Numbering(new Map(), new Map());
    }

    static parse(source: string): Numbering {
        const root = XmlDocument.parse(source).root;

        const abstracts = new Map<string, AbstractNumbering>();
        for (const abstract of root.elements('w:abstractNum')) {
            const id = abstract.attribute('w:abstractNumId');
            if (null === id) {
                continue;
            }

            const levels = new Map<number, NumberingLevel>();
            for (const level of abstract.elements('w:lvl')) {
                const ilvl = number(level.attribute('w:ilvl'));
                if (null !== ilvl) {
                    levels.set(ilvl, readLevel(level));
                }
            }

            abstracts.set(id, { levels });
        }

        const instances = new Map<string, NumberingInstance>();
        for (const instance of root.elements('w:num')) {
            const numId = instance.attribute('w:numId');
            const abstractId = instance.element('w:abstractNumId')?.attribute('w:val') ?? null;
            if (null === numId || null === abstractId) {
                continue;
            }

            const overrides = new Map<number, LevelOverride>();
            for (const override of instance.elements('w:lvlOverride')) {
                const ilvl = number(override.attribute('w:ilvl'));
                if (null === ilvl) {
                    continue;
                }

                const replacement = override.element('w:lvl') ?? null;
                overrides.set(ilvl, {
                    startAt: number(
                        override.element('w:startOverride')?.attribute('w:val') ?? null),
                    level: null === replacement ? null : readLevel(replacement),
                });
            }

            instances.set(numId, { abstractId, overrides });
        }

        return new Numbering(abstracts, instances);
    }

    level(numId: string, ilvl: number): NumberingLevel | null {
        const instance = this.instances.get(numId);
        if (undefined === instance) {
            return null;
        }

        // The instance's own level wins outright where it states one: it is a
        // replacement rather than a patch, which is why nothing is merged here.
        return instance.overrides.get(ilvl)?.level
            ?? this.abstracts.get(instance.abstractId)?.levels.get(ilvl)
            ?? null;
    }

    /**
     * Which definition this numId counts on.
     *
     * The counters key on this rather than on the numId, because that is where
     * the running number lives — see the note at the top of this file.
     */
    abstractOf(numId: string): string | null {
        return this.instances.get(numId)?.abstractId ?? null;
    }

    /** `w:startOverride`, if this instance restarts the count at this level. */
    startOverride(numId: string, ilvl: number): number | null {
        return this.instances.get(numId)?.overrides.get(ilvl)?.startAt ?? null;
    }
}

/** One `w:lvl`, wherever it is written — in an abstract or in an override. */
function readLevel(level: XmlElement): NumberingLevel {
    const indent = level.element('w:pPr')?.element('w:ind') ?? null;

    return {
        format: toFormat(level.element('w:numFmt')?.attribute('w:val') ?? null),
        text: level.element('w:lvlText')?.attribute('w:val') ?? '',
        start: number(level.element('w:start')?.attribute('w:val') ?? null) ?? 1,
        indentLeftTwips: number(indent?.attribute('w:left') ?? indent?.attribute('w:start') ?? null),
        indentHangingTwips: number(indent?.attribute('w:hanging') ?? null),
        fontFamily: level.element('w:rPr')?.element('w:rFonts')?.attribute('w:ascii') ?? null,
        suffix: toSuffix(level.element('w:suff')?.attribute('w:val') ?? null),
        justification: toJustification(level.element('w:lvlJc')?.attribute('w:val') ?? null),
        restart: number(level.element('w:lvlRestart')?.attribute('w:val') ?? null),
    };
}

/**
 * The running counters for a document's lists.
 *
 * Kept outside {@link Numbering} because the definitions are shared and the
 * counts are not: reading the same document twice must produce the same
 * numbers, and it would not if the counters lived on the definition.
 */
export class NumberingCounters {
    /** abstractNumId to level to current value. */
    private readonly counts = new Map<string, Map<number, number>>();

    /**
     * The `numId:ilvl` pairs whose `w:startOverride` has already fired.
     *
     * An override restarts the count ONCE, at the first paragraph using its
     * numId, and is spent from then on — measured, where a list
     * resumed at 11 and 12 rather than restarting at its stated 7.
     */
    private readonly spent = new Set<string>();

    /**
     * Advance the counter for one list paragraph and render its marker.
     *
     * Deeper levels RESET when a shallower one advances — that is what makes
     * the second "1.1" start at one again rather than continuing from the
     * first section's count.
     */
    next(numbering: Numbering, numId: string, ilvl: number): string {
        const level = numbering.level(numId, ilvl);
        if (null === level) {
            return '';
        }
        if ('bullet' === level.format || 'none' === level.format) {
            // A bullet has no count: its lvlText IS the glyph.
            return level.text;
        }

        // On the ABSTRACT, so a document that changes numId part-way through
        // one list — which is what Word writes for a list whose formatting
        // changes — carries on counting rather than starting again.
        const key = numbering.abstractOf(numId) ?? numId;
        const counts = this.counts.get(key) ?? new Map<number, number>();
        this.counts.set(key, counts);

        const restartAt = numbering.startOverride(numId, ilvl);
        if (null !== restartAt && !this.spent.has(`${numId}:${ilvl}`)) {
            this.spent.add(`${numId}:${ilvl}`);
            counts.set(ilvl, restartAt - 1);
        }

        counts.set(ilvl, (counts.get(ilvl) ?? level.start - 1) + 1);
        for (const deeper of [...counts.keys()]) {
            // A level with w:lvlRestart="0" keeps counting: it is how a
            // document numbers its figures 1..40 across chapters that each
            // restart their own sections.
            //
            // THE REFERENCE DOES NOT DO THIS. LibreOffice ignores
            // `w:lvlRestart` outright: the same three-level list printed
            // `i. ii. b) i.` whether the level said nothing, said `0`, or said
            // `1`. We follow the FILE — as we do with the VML inset — and for
            // a second reason here: Word honours the element (it is the
            // "restart numbering after" setting), so a document authored there
            // reads correctly this way and not the other.
            //
            // A value ABOVE zero names the level whose change restarts this
            // one, and is not read: anything non-zero restarts under any
            // shallower level. It cannot be measured against a renderer that
            // ignores the attribute, so `readNumbering` reports it instead.
            if (deeper > ilvl && 0 !== numbering.level(numId, deeper)?.restart) {
                counts.delete(deeper);
            }
        }

        return render(level, numbering, numId, counts);
    }
}

/**
 * Substitute `%1`, `%2` … with the counters of levels one, two and so on.
 *
 * A pattern refers to LEVELS, not to itself, which is how "1.2.3" is written
 * once at the third level and picks up the two above it.
 */
function render(
    level: NumberingLevel,
    numbering: Numbering,
    numId: string,
    counts: ReadonlyMap<number, number>,
): string {
    return level.text.replace(/%(\d+)/g, (_match, digits: string) => {
        const referenced = Number(digits) - 1;
        const value = counts.get(referenced)
            ?? numbering.level(numId, referenced)?.start
            ?? 1;

        return format(value, numbering.level(numId, referenced)?.format ?? level.format);
    });
}

/**
 * A list's number, written the way its level asks for.
 *
 * `bullet` and `none` are not numbers at all and fall through to the digits —
 * a level of either draws its `w:lvlText` verbatim and never reaches here with
 * a value that matters.
 */
function format(value: number, style: NumberFormat): string {
    return 'bullet' === style || 'none' === style
        ? String(value)
        : formatNumeral(value, style);
}

function toSuffix(value: string | null): 'tab' | 'space' | 'nothing' {
    return 'space' === value || 'nothing' === value ? value : 'tab';
}

/**
 * `w:lvlJc`, which the strict schema spells `start`/`end` as `w:jc` does.
 *
 * Anything else — including the absent case — is `left`, which is what every
 * list in the corpus states and what Word writes for a new one.
 */
function toJustification(value: string | null): 'left' | 'right' | 'center' {
    switch (value) {
        case 'right':
        case 'end':
            return 'right';
        case 'center':
            return 'center';
        default:
            return 'left';
    }
}

function toFormat(value: string | null): NumberFormat {
    switch (value) {
        case 'bullet':
        case 'decimal':
        case 'lowerLetter':
        case 'upperLetter':
        case 'lowerRoman':
        case 'upperRoman':
        case 'none':
            return value;
        default:
            // Everything else — ordinal, cardinalText, chicago, the CJK formats
            // — counts like decimal here. Wrong glyphs are better than a list
            // that silently loses its numbers.
            return null === value ? 'none' : 'decimal';
    }
}

function number(value: string | null): number | null {
    if (null === value) {
        return null;
    }
    const parsed = Number(value);

    return Number.isFinite(parsed) ? parsed : null;
}
