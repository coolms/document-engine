import { OpcPackage, type ZipCodec } from '../opc/opc-package.js';
import {
    OFFICE_DOCUMENT,
    findByType,
    readRelationships,
    relationshipPartFor,
    resolveTarget,
} from '../opc/relationships.js';
import { readWordDocument, type WordDocument } from './document-reader.js';
import type { ImageContent } from '../layout/image.js';
import type { FontCatalogue } from './font-catalogue.js';

/**
 * Opening a `.docx` FILE — the last link in the chain.
 *
 * Bytes in, a paginated document out, and the package still open so the same
 * file can be written back with everything untouched preserved.
 */

/** A `.docx` opened: the laid-out document, and the package it came from. */
export interface OpenedWordFile {
    readonly document: WordDocument;
    /** Still open, so an edit can be written back into the original container. */
    readonly package: OpcPackage;
    /** The main part's name, whatever the package chose to call it. */
    readonly documentPart: string;
    readonly stylesPart: string | null;
    readonly numberingPart: string | null;
}

export async function openWordFile(
    bytes: Uint8Array,
    fonts: FontCatalogue,
    codec?: ZipCodec,
): Promise<OpenedWordFile> {
    const opc = undefined === codec ? OpcPackage.open(bytes) : OpcPackage.open(bytes, codec);

    const rootRelationships = await opc.textIfPresent('_rels/.rels');
    if (null === rootRelationships) {
        throw new Error('Not an OOXML package: it has no _rels/.rels.');
    }

    const main = findByType(readRelationships(rootRelationships), OFFICE_DOCUMENT);
    if (null === main) {
        throw new Error('Package declares no officeDocument relationship; it is not a Word document.');
    }

    const documentPart = resolveTarget('', main.target);
    if (!opc.has(documentPart)) {
        throw new Error(`Package points at "${documentPart}" as its main part, but has no such part.`);
    }

    const stylesPart = await related(opc, documentPart, 'styles');
    const numberingPart = await related(opc, documentPart, 'numbering');
    const settingsPart = await related(opc, documentPart, 'settings');
    const footnotesPart = await related(opc, documentPart, 'footnotes');
    const endnotesPart = await related(opc, documentPart, 'endnotes');

    const stylesXml = null === stylesPart ? null : await opc.text(stylesPart);
    const numberingXml = null === numberingPart ? null : await opc.text(numberingPart);
    const settingsXml = null === settingsPart ? null : await opc.text(settingsPart);
    const footnotesXml = null === footnotesPart ? null : await opc.text(footnotesPart);
    const endnotesXml = null === endnotesPart ? null : await opc.text(endnotesPart);

    const documentXml = await opc.text(documentPart);
    const furniture = await readFurniture(opc, documentPart);
    const mediaById = await readMedia(opc, documentPart);

    const document = readWordDocument({
        documentXml,
        fonts,
        ...(null === stylesXml ? {} : { stylesXml }),
        ...(null === numberingXml ? {} : { numberingXml }),
        ...(null === settingsXml ? {} : { settingsXml }),
        ...(null === footnotesXml ? {} : { footnotesXml }),
        ...(null === endnotesXml ? {} : { endnotesXml }),
        ...furniture,
        ...(0 === Object.keys(mediaById).length ? {} : { mediaById }),
    });

    return { document, package: opc, documentPart, stylesPart, numberingPart };
}

/**
 * Every header and footer part in the package, keyed by relationship id.
 *
 * Not by variant, and not per section: a `w:headerReference` names an `r:id`,
 * and only the reader — which is already walking the section properties — knows
 * which section names which. Working it out here would mean splitting the body a
 * second time, and the two would eventually disagree.
 */
async function readFurniture(
    opc: OpcPackage,
    documentPart: string,
): Promise<{ furnitureById?: Record<string, string> }> {
    const relationshipsXml = await opc.textIfPresent(relationshipPartFor(documentPart));
    if (null === relationshipsXml) {
        return {};
    }

    const furnitureById: Record<string, string> = {};
    for (const relationship of readRelationships(relationshipsXml)) {
        const type = relationship.type.split('/').pop();
        if (relationship.external || ('header' !== type && 'footer' !== type)) {
            continue;
        }

        const xml = await opc.textIfPresent(resolveTarget(documentPart, relationship.target));
        if (null !== xml) {
            furnitureById[relationship.id] = xml;
        }
    }

    return 0 === Object.keys(furnitureById).length ? {} : { furnitureById };
}

/**
 * Every picture the main document points at, keyed by relationship id.
 *
 * Read eagerly, because a drawing names its picture by `r:embed` and the reader
 * is synchronous — fetching a part while walking the body would make the whole
 * read asynchronous for the sake of one element type.
 *
 * An EXTERNAL relationship is skipped: its target is a URL, and fetching it
 * would turn opening a document into a network request the caller never asked
 * for.
 */
async function readMedia(
    opc: OpcPackage,
    documentPart: string,
): Promise<Record<string, ImageContent>> {
    const relationshipsXml = await opc.textIfPresent(relationshipPartFor(documentPart));
    if (null === relationshipsXml) {
        return {};
    }

    const media: Record<string, ImageContent> = {};
    for (const relationship of readRelationships(relationshipsXml)) {
        if (relationship.external || 'image' !== relationship.type.split('/').pop()) {
            continue;
        }

        const part = resolveTarget(documentPart, relationship.target);
        if (!opc.has(part)) {
            continue;
        }

        media[relationship.id] = {
            bytes: await opc.bytes(part),
            contentType: contentTypeOf(part),
        };
    }

    return media;
}

/**
 * A picture's media type, from its extension.
 *
 * From the extension rather than from `[Content_Types].xml` because the two
 * agree for every picture Word writes, and the alternative is threading the
 * content-type part through a reader that needs nothing else from it. An
 * unknown extension gets `application/octet-stream`, which a renderer declines
 * to draw rather than mislabels.
 */
function contentTypeOf(part: string): string {
    const extension = part.slice(part.lastIndexOf('.') + 1).toLowerCase();

    return IMAGE_TYPES[extension] ?? 'application/octet-stream';
}

const IMAGE_TYPES: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    svg: 'image/svg+xml',
    webp: 'image/webp',
    emf: 'image/x-emf',
    wmf: 'image/x-wmf',
};

/**
 * A part the main document points at, by relationship type.
 *
 * A document without styles or without numbering is legal — it simply has
 * none — so this returns null rather than throwing.
 */
async function related(opc: OpcPackage, documentPart: string, type: string): Promise<string | null> {
    const xml = await opc.textIfPresent(relationshipPartFor(documentPart));
    if (null === xml) {
        return null;
    }

    const relationship = findByType(readRelationships(xml), type);
    if (null === relationship) {
        return null;
    }

    const part = resolveTarget(documentPart, relationship.target);

    return opc.has(part) ? part : null;
}
