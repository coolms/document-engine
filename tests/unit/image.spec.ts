import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { TrueTypeFont } from '../../src/font/truetype-font.js';
import { OBJECT_REPLACEMENT, type InlineImage } from '../../src/layout/image.js';
import { breakIntoLines } from '../../src/layout/line-breaker.js';
import {
    DEFAULT_PARAGRAPH_STYLE,
    layoutPages,
    type PageGeometry,
    type Paragraph,
} from '../../src/layout/page-layout.js';
import { XmlDocument } from '../../src/ooxml/xml.js';
import { renderPage, type ImageOp } from '../../src/render/display-list.js';
import { renderPageToSvg } from '../../src/render/svg.js';

const FONT_DIR = join(dirname(fileURLToPath(import.meta.url)), '../../assets/fonts');
const MONO = TrueTypeFont.parse(new Uint8Array(readFileSync(join(FONT_DIR, 'LiberationMono-Regular.ttf'))));

const CELL = MONO.measureAdvance('x', 16).widthPt;

const PAGE: PageGeometry = {
    widthPx: 600,
    heightPx: 600,
    marginTopPx: 10,
    marginRightPx: 10,
    marginBottomPx: 10,
    marginLeftPx: 10,
};

const BYTES = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 250, 251, 252]);

const image = (widthPx: number, heightPx: number, description?: string): InlineImage => ({
    content: {
        bytes: BYTES,
        contentType: 'image/png',
        ...(undefined === description ? {} : { description }),
    },
    widthPx,
    heightPx,
});

const picture = (widthPx: number, heightPx: number, description?: string): Paragraph => ({
    runs: [
        { text: 'ab', font: MONO, sizePx: 16 },
        { text: OBJECT_REPLACEMENT, font: MONO, sizePx: 16, image: image(widthPx, heightPx, description) },
        { text: 'cd', font: MONO, sizePx: 16 },
    ],
    style: DEFAULT_PARAGRAPH_STYLE,
});

const imageOps = (paragraph: Paragraph): ImageOp[] =>
    renderPage(layoutPages([paragraph], PAGE)[0]!).ops
        .filter((op): op is ImageOp => 'image' === op.kind);

describe('an inline picture', () => {
    it('takes its width from the DOCUMENT, not from measuring its stand-in glyph', () => {
        // The run's text is one object-replacement character, which has no
        // glyph worth the name — measuring it would give whatever the font
        // keeps at .notdef.
        const [line] = breakIntoLines(
            [{ text: OBJECT_REPLACEMENT, font: MONO, sizePx: 16, image: image(120, 60) }],
            1000,
        );

        expect(line!.pieces[0]!.widthPx).toBe(120);
        expect(line!.widthPx).toBe(120);
    });

    it('still occupies exactly one unit of source', () => {
        // So a caret can be put either side of it and every offset after it
        // stays where the document says.
        const [line] = breakIntoLines(
            [{ text: OBJECT_REPLACEMENT, font: MONO, sizePx: 16, image: image(120, 60) }],
            1000,
        );

        expect(line!.endsAt - line!.startsAt).toBe(1);
    });

    it('makes its line tall enough to hold it', () => {
        const [page] = layoutPages([picture(120, 200)], PAGE);
        const line = page!.lines[0]!;

        // The picture stands ON the baseline, so the line holds all of it above
        // the baseline and the font's descender below.
        expect(line.baselinePx).toBeCloseTo(200, 6);
        expect(line.heightPx).toBeCloseTo(200 + MONO.descent(16), 6);
    });

    it('leaves a short picture to the text around it', () => {
        // A picture smaller than the line does not SHRINK it.
        const [page] = layoutPages([picture(20, 4)], PAGE);

        expect(page!.lines[0]!.heightPx).toBeCloseTo(MONO.naturalLineHeight(16), 6);
    });

    it('stands on the baseline, and the text carries on after it', () => {
        const [page] = layoutPages([picture(120, 60)], PAGE);
        const line = page!.lines[0]!;
        const [op] = imageOps(picture(120, 60));

        // Drawn by its TOP-LEFT, a whole height above the baseline.
        expect(op!.yPx + op!.heightPx).toBeCloseTo(line.yPx + line.baselinePx, 6);
        // Two characters of text before it, and its own width after.
        expect(op!.xPx).toBeCloseTo(10 + CELL * 2, 6);

        const after = renderPage(layoutPages([picture(120, 60)], PAGE)[0]!).ops
            .filter((each) => 'text' === each.kind);
        expect(after[after.length - 1]!.xPx).toBeCloseTo(10 + CELL * 2 + 120, 6);
    });

    describe('serialised', () => {
        const svgOf = (paragraph: Paragraph): string =>
            renderPageToSvg(layoutPages([paragraph], PAGE)[0]!);

        it('embeds the bytes rather than pointing at a part nobody has', () => {
            const element = XmlDocument.parse(svgOf(picture(120, 60))).root.element('image');
            const href = element?.attribute('href') ?? '';

            expect(href.startsWith('data:image/png;base64,')).toBe(true);
            // Decoded back, byte for byte — a truncated tail would still look
            // like a plausible data URI.
            const decoded = Buffer.from(href.slice('data:image/png;base64,'.length), 'base64');
            expect(new Uint8Array(decoded)).toEqual(BYTES);
        });

        it('pads base64 so a STRICT decoder does not lose the tail', () => {
            // Node's decoder is lenient and accepts unpadded input, so a
            // round-trip through it cannot see missing padding at all. The
            // length is what a strict decoder actually requires: base64 comes
            // in groups of four characters, always.
            const lengths = [1, 2, 3, 4, 5];
            for (const length of lengths) {
                const bytes = BYTES.slice(0, length);
                const paragraph: Paragraph = {
                    runs: [{
                        text: OBJECT_REPLACEMENT,
                        font: MONO,
                        sizePx: 16,
                        image: { content: { bytes, contentType: 'image/png' }, widthPx: 10, heightPx: 10 },
                    }],
                    style: DEFAULT_PARAGRAPH_STYLE,
                };
                const href = XmlDocument.parse(svgOf(paragraph)).root.element('image')
                    ?.attribute('href') ?? '';
                const payload = href.slice('data:image/png;base64,'.length);
                const decoded = Buffer.from(payload, 'base64');

                expect(payload.length % 4).toBe(0);
                expect(new Uint8Array(decoded)).toEqual(bytes);
            }
        });

        it('fits the picture to its box instead of stretching it', () => {
            // A document's aspect ratio is not the renderer's to change.
            expect(XmlDocument.parse(svgOf(picture(120, 60))).root.element('image')
                ?.attribute('preserveAspectRatio')).toBe('xMidYMid meet');
        });

        it('carries the alternative text through when there is one', () => {
            const withText = XmlDocument.parse(svgOf(picture(120, 60, 'A red rectangle')))
                .root.element('image');
            const without = XmlDocument.parse(svgOf(picture(120, 60))).root.element('image');

            expect(withText?.attribute('aria-label')).toBe('A red rectangle');
            expect(without?.attribute('aria-label')).toBeNull();
        });
    });
});
