/**
 * The corpus the engine's arithmetic is checked against a REAL text engine.
 *
 * The engine deliberately never asks the DOM how wide text is — it reads the
 * font. That removes browser-vs-renderer divergence by construction, but it
 * moves the risk somewhere else: if our reading of the font is wrong, every
 * page boundary is wrong and nothing else in the system would notice.
 *
 * So these cases are measured both ways. Kerning-free strings must agree
 * EXACTLY; a known kerned pair is included precisely because it must NOT agree
 * yet, which measures the size of the gap GPOS support has to close instead of
 * leaving it as a vague "kerning is todo".
 */
export interface ParityCase {
    readonly file: string;
    /** CSS family name the browser side registers the file under. */
    readonly family: string;
    readonly text: string;
    readonly sizePx: number;
    /**
     * Whether the two engines are expected to agree today. `false` marks a
     * measured, deliberate divergence — not a failure.
     */
    readonly kerningFree: boolean;
    readonly note: string;
}

export const PARITY_CASES: readonly ParityCase[] = [
    {
        file: 'LiberationMono-Regular.ttf',
        family: 'ParityMono',
        text: 'xxxxxxxxxx',
        sizePx: 16,
        kerningFree: true,
        note: 'Monospaced: every advance identical, so no engine can disagree.',
    },
    {
        file: 'LiberationMono-Regular.ttf',
        family: 'ParityMono',
        text: 'Договор аренды',
        sizePx: 16,
        kerningFree: true,
        note: 'Cyrillic through a monospaced face — the documents in use are Russian.',
    },
    {
        file: 'Caladea-Regular.ttf',
        family: 'ParitySerif1000',
        text: 'nnnnmmmm',
        sizePx: 16,
        kerningFree: true,
        note: 'Proportional with unitsPerEm 1000, and no kerned pairs among n/m.',
    },
    {
        file: 'LiberationSerif-Regular.ttf',
        family: 'ParitySerif2048',
        text: 'nonsense',
        sizePx: 16,
        kerningFree: true,
        note: 'Proportional with unitsPerEm 2048.',
    },
    {
        file: 'LiberationSerif-Regular.ttf',
        family: 'ParitySerif2048',
        text: 'AVAVAVAV',
        sizePx: 16,
        kerningFree: true,
        note: 'AV is the classic kerned pair, and we kern it too. This case '
            + 'measured the GPOS gap before that was built (14.4375px over four '
            + 'pairs) and now guards it: a regression in GPOS reappears here.',
    },
    {
        file: 'LiberationMono-Regular.ttf',
        family: 'ParityMono',
        text: 'AVAVAVAV',
        sizePx: 16,
        kerningFree: true,
        note: 'The SAME pairs in a monospaced face, which has no kern feature in '
            + 'its GPOS at all — only mark. Kerning AV here would break the grid '
            + 'that makes the font monospaced, so this guards against applying '
            + 'adjustments a font does not actually declare.',
    },
    {
        file: 'Carlito-Regular.ttf',
        family: 'ParityDefault',
        text: 'Wave To Avoid',
        sizePx: 16,
        kerningFree: true,
        note: 'The DEFAULT font, with several classic kerned pairs at once. '
            + 'Carlito has no legacy kern table, so this case only passes if '
            + 'GPOS is being read.',
    },
];
