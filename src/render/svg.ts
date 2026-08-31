import { escapeAttribute, escapeText } from '../ooxml/xml.js';
import type { Page } from '../layout/page-layout.js';
import { renderPage, type ImageOp, type LineOp, type RectOp, type RenderedPage, type TextOp }
    from './display-list.js';

/**
 * One page as an SVG document.
 *
 * SVG because it is a STRING: it can be asserted on in a test, diffed between
 * runs, printed, and handed to a browser without a canvas or a DOM. The display
 * list beside it holds the geometry, so this file only serialises — anything
 * here that had to calculate a position would be a second layout engine.
 *
 * ## Fonts are referenced, not embedded
 *
 * Each `<text>` names the family the FILE declares, so the SVG asks for the
 * face the engine actually measured with rather than the one the document
 * requested. Whether that face is available is the viewer's business: pass
 * `styleCss` with `@font-face` rules to make the page self-contained.
 */

export interface SvgOptions {
    /**
     * Extra CSS for the page — `@font-face` rules, most usefully.
     *
     * Emitted inside a CDATA section, so it may contain the `&` and `<` that a
     * `url(...)` or a media query brings with it.
     */
    readonly styleCss?: string;
    /** Paper colour. Null draws no background at all, for compositing. */
    readonly backgroundColor?: string | null;
    /**
     * Colour for text that does not name one of its own.
     *
     * A run's own `w:color` always wins; this is the page's default, which is
     * what a document drawn light-on-dark needs to change.
     */
    readonly textColor?: string;
}

export function renderPageToSvg(page: Page, options: SvgOptions = {}): string {
    return serialise(renderPage(page), options);
}

export function serialise(rendered: RenderedPage, options: SvgOptions = {}): string {
    const background = undefined === options.backgroundColor ? '#ffffff' : options.backgroundColor;
    const parts: string[] = [];

    parts.push(
        '<svg xmlns="http://www.w3.org/2000/svg"'
        + ` width="${round(rendered.widthPx)}" height="${round(rendered.heightPx)}"`
        + ` viewBox="0 0 ${round(rendered.widthPx)} ${round(rendered.heightPx)}">`,
    );

    if (undefined !== options.styleCss) {
        parts.push(`<style><![CDATA[${options.styleCss}]]></style>`);
    }

    if (null !== background) {
        parts.push(
            `<rect width="${round(rendered.widthPx)}" height="${round(rendered.heightPx)}"`
            + ` fill="${escapeAttribute(background)}"/>`,
        );
    }

    for (const op of rendered.ops) {
        if ('rect' === op.kind) {
            parts.push(rect(op));
        } else if ('line' === op.kind) {
            parts.push(line(op));
        } else if ('image' === op.kind) {
            parts.push(picture(op));
        } else {
            parts.push(text(op, op.colorHex ?? options.textColor ?? '#000000'));
        }
    }

    parts.push('</svg>');

    return parts.join('');
}

/**
 * A picture, embedded rather than linked.
 *
 * Inlined as a data URI because the page has to stay ONE string: an `<image>`
 * pointing at `media/image1.png` would resolve against whatever directory the
 * SVG was opened from, and the bytes are inside a zip the viewer does not have.
 */
function picture(op: ImageOp): string {
    const source = `data:${op.image.content.contentType};base64,${base64(op.image.content.bytes)}`;

    return `<image x="${round(op.xPx)}" y="${round(op.yPx)}"`
        + ` width="${round(op.widthPx)}" height="${round(op.heightPx)}"`
        + (undefined === op.image.content.description
            ? ''
            : ` aria-label="${escapeAttribute(op.image.content.description)}"`)
        // Without this the picture is stretched to fill the box rather than
        // fitted to it, and a document's aspect ratio is not ours to change.
        + ' preserveAspectRatio="xMidYMid meet"'
        + ` href="${source}"/>`;
}

const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/**
 * Base64, written out rather than taken from the platform.
 *
 * `Buffer` is Node's and `btoa` is the browser's, and this package runs in
 * both — an editor in the browser, a renderer on the server. Reaching for
 * either would fail at RUN time in the other rather than at build time.
 */
function base64(bytes: Uint8Array): string {
    let out = '';

    for (let index = 0; index < bytes.length; index += 3) {
        const a = bytes[index]!;
        const b = bytes[index + 1];
        const c = bytes[index + 2];
        const triple = (a << 16) | ((b ?? 0) << 8) | (c ?? 0);

        out += BASE64_ALPHABET[(triple >> 18) & 63];
        out += BASE64_ALPHABET[(triple >> 12) & 63];
        // The tail is PADDED, not truncated: a decoder takes three bytes from
        // every four characters, and a short final group shifts nothing but
        // silently loses the last byte or two.
        out += undefined === b ? '=' : BASE64_ALPHABET[(triple >> 6) & 63];
        out += undefined === c ? '=' : BASE64_ALPHABET[triple & 63];
    }

    return out;
}

function rect(op: RectOp): string {
    return `<rect x="${round(op.xPx)}" y="${round(op.yPx)}"`
        + ` width="${round(op.widthPx)}" height="${round(op.heightPx)}"`
        + ` fill="${escapeAttribute(op.fill)}"/>`;
}

/**
 * A border, as a stroke.
 *
 * `double` is drawn as a single rule of the declared width rather than as two
 * thinner ones: the width is what the layout reserved, and drawing two strokes
 * inside it would need a gap the model does not record. Stated rather than
 * approximated, since a double border drawn as two hairlines is further from
 * the document than one solid rule of the right weight.
 */
function line(op: LineOp): string {
    const dash = 'dashed' === op.style
        ? ` stroke-dasharray="${round(op.widthPx * 3)},${round(op.widthPx * 2)}"`
        : ('dotted' === op.style ? ` stroke-dasharray="${round(op.widthPx)},${round(op.widthPx * 2)}"` : '');

    return `<line x1="${round(op.x1Px)}" y1="${round(op.y1Px)}"`
        + ` x2="${round(op.x2Px)}" y2="${round(op.y2Px)}"`
        + ` stroke="${escapeAttribute(op.color)}" stroke-width="${round(op.widthPx)}"${dash}/>`;
}

function text(op: TextOp, colour: string): string {
    const family = op.font.familyName;
    const subfamily = op.font.subfamilyName.toLowerCase();

    // Weight and slant live in separate FILES here, so they cannot be read off
    // the run — only off the face that was chosen for it.
    const weight = subfamily.includes('bold') ? ' font-weight="bold"' : '';
    const slant = subfamily.includes('italic') || subfamily.includes('oblique')
        ? ' font-style="italic"'
        : '';

    // A quarter turn about the operation's OWN origin, which is its baseline
    // start — the same point the PDF's text matrix turns about, so the two
    // renderers put a turned run in the same place.
    const turn = undefined === op.turn
        ? ''
        : ` transform="rotate(${'ccw' === op.turn ? -90 : 90} ${round(op.xPx)}`
            + ` ${round(op.yPx)})"`;

    return `<text x="${round(op.xPx)}" y="${round(op.yPx)}"${turn}`
        + ('' === family ? '' : ` font-family="${escapeAttribute(family)}"`)
        + ` font-size="${round(op.sizePx)}"${weight}${slant}`
        + ` fill="${escapeAttribute(colour)}"`
        // Stated both ways round rather than left to the viewer's default,
        // which kerns: a string measured unkerned and then drawn kerned puts
        // its later glyphs left of where the line breaker put them.
        // As a style rather than a presentation attribute because `font-kerning`
        // only became one in SVG 2, and the CSS property is honoured either way.
        + ` style="font-kerning:${true === op.kerned ? 'normal' : 'none'}"`
        // Without this a leading or trailing space in a piece is collapsed away
        // and the words after it shift left of where they were measured.
        + ` xml:space="preserve">${escapeText(op.text)}</text>`;
}

/**
 * Positions to three decimals.
 *
 * Enough that a rounding error cannot reach a device pixel at any sane zoom,
 * and short enough that the output stays diffable between runs.
 */
function round(value: number): string {
    return String(Math.round(value * 1000) / 1000);
}
