/**
 * Where a line is ALLOWED to break.
 *
 * A deliberate subset of UAX #14, not an implementation of it. The full
 * algorithm is a pair table over ~40 character classes; what document text
 * actually needs is: break after spaces, break after hyphens, break between
 * ideographs, and never break a word otherwise. Everything here is a rule that
 * earns its place by changing where real text breaks.
 *
 * What is deliberately NOT here:
 *
 *   - **Hyphenation.** Splitting `international` into `inter-national` needs a
 *     per-language dictionary and is a separate feature with its own switch.
 *     Without one, an over-long word overflows rather than being guessed at.
 *   - **Line-break tailoring per locale** (French spacing before `!`, Japanese
 *     kinsoku). Both change where a break may fall and both belong with the
 *     language data, once the model carries a language.
 */

/** Space characters after which a line may break. */
const BREAK_AFTER_SPACE = new Set([
    0x0020, // SPACE
    0x0009, // TAB
    0x1680, // OGHAM SPACE MARK
    0x2000, 0x2001, 0x2002, 0x2003, 0x2004, 0x2005, 0x2006, // EN QUAD … SIX-PER-EM
    0x2008, 0x2009, 0x200a, // PUNCTUATION SPACE, THIN, HAIR
    0x205f, // MEDIUM MATHEMATICAL SPACE
    0x3000, // IDEOGRAPHIC SPACE
]);

/**
 * Spaces that explicitly REFUSE a break. The whole reason they exist.
 *
 * U+00A0 in `10 kg` or `Mr. Smith` is the author saying "not here", and a
 * breaker that treats it as an ordinary space silently overrides them.
 */
const NO_BREAK_SPACE = new Set([
    0x00a0, // NO-BREAK SPACE
    0x2007, // FIGURE SPACE
    0x202f, // NARROW NO-BREAK SPACE
]);

/** Characters after which a break is allowed even mid-word. */
const BREAK_AFTER = new Set([
    0x002d, // HYPHEN-MINUS
    0x00ad, // SOFT HYPHEN — the author's own offer of a break
    0x2010, // HYPHEN
    0x2012, 0x2013, 0x2014, // FIGURE DASH, EN DASH, EM DASH
    0x200b, // ZERO WIDTH SPACE
]);

/**
 * Characters before which a break must NOT happen, whatever precedes them.
 *
 * The NON-BREAKING HYPHEN is NOT among them, though Unicode gives it a class
 * that forbids a break on either side. Measured: a line reading
 * `www ‑www`, too long for its column, broke at the SPACE and carried the
 * hyphen to the head of the next line. Forbidding that break made the whole
 * fragment move and then chopped it mid-word, which is worse and is not what
 * the page shows.
 *
 * So the character refuses a break AFTER itself — it is absent from
 * {@link BREAK_AFTER}, which is the whole of its job — and says nothing about
 * what comes before.
 */
const NO_BREAK_BEFORE = new Set([
    0x2060, // WORD JOINER
]);

export function isSpace(codePoint: number): boolean {
    return BREAK_AFTER_SPACE.has(codePoint) || NO_BREAK_SPACE.has(codePoint);
}

/**
 * The SOFT HYPHEN, which is a break the author offered and no ink until it is
 * taken.
 *
 * Measured against LibreOffice: a word split by one broke AT it and
 * printed a `-` at the end of that line, where the same word without one broke
 * where the width ran out and printed nothing. So it costs nothing mid-line and
 * costs a hyphen at the end of one — which is the difference between a
 * segment's `full` width and its `content` width, and why it needs no new
 * notion in the breaker.
 */
export const SOFT_HYPHEN = 0x00ad;

/**
 * Where a break is FORBIDDEN outright, and not merely unoffered.
 *
 * The difference matters only to a chop: a word too long for its own
 * line is cut wherever the width runs out, mid-letter and against every
 * ordinary rule, because the alternative is text running off the page. These
 * two are the cases where the author said "not here" in so many words, and a
 * chop has to go round them the way an ordinary break does — `10 kg` is
 * kept whole by U+00A0 whether the line is short of room or not.
 */
export function forbidsBreakBetween(before: number, after: number): boolean {
    // Both SIDES of a no-break space: cutting in front of one would start the
    // next line with it, which is the same join broken from the other end.
    return NO_BREAK_SPACE.has(before)
        || NO_BREAK_SPACE.has(after)
        || NO_BREAK_BEFORE.has(after);
}

export function isBreakableSpace(codePoint: number): boolean {
    return BREAK_AFTER_SPACE.has(codePoint);
}

/** A hard break the author typed: always ends the line. */
export function isMandatoryBreak(codePoint: number): boolean {
    return 0x000a === codePoint // LINE FEED
        || 0x000d === codePoint // CARRIAGE RETURN
        || 0x2028 === codePoint // LINE SEPARATOR
        || 0x2029 === codePoint; // PARAGRAPH SEPARATOR
}

/**
 * Ideographs and kana, which break between almost any two characters.
 *
 * Included because a CJK paragraph with no break opportunities at all would be
 * measured as one enormous unbreakable word and overflow the page — a failure
 * severe enough to be worth the small table even before the language layer
 * exists.
 */
export function isIdeographic(codePoint: number): boolean {
    return (codePoint >= 0x3040 && codePoint <= 0x30ff) // Hiragana, Katakana
        || (codePoint >= 0x3400 && codePoint <= 0x4dbf) // CJK Ext A
        || (codePoint >= 0x4e00 && codePoint <= 0x9fff) // CJK Unified
        || (codePoint >= 0xf900 && codePoint <= 0xfaff); // Compatibility
}

/**
 * May a line break BETWEEN these two code points?
 *
 * `before` is the character ending the current line, `after` the one starting
 * the next. Both are needed: a break is a property of the pair, not of either
 * character — `-` allows one after it, but not when a word joiner follows.
 */
export function allowsBreakBetween(before: number, after: number): boolean {
    if (NO_BREAK_BEFORE.has(after)) {
        return false;
    }

    // NOTE: there is deliberately no `NO_BREAK_SPACE.has(before)` check here.
    // A non-breaking space is protected by ABSENCE — it is in no set that
    // grants a break — so an explicit rejection would be dead code that reads
    // as load-bearing. Mutation testing caught exactly that: disabling such a
    // check changed no behaviour and broke no test.
    if (isBreakableSpace(before)) {
        // A run of spaces breaks after the LAST of them, not between them, so
        // the trailing whitespace stays with the line it ends.
        return !isSpace(after);
    }
    if (BREAK_AFTER.has(before)) {
        return true;
    }

    return isIdeographic(before) && isIdeographic(after);
}
