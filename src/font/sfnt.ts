/**
 * Minimal sfnt (TrueType / OpenType) container reader.
 *
 * ## Why we parse fonts at all
 *
 * The layout engine must not ask the DOM how wide text is. If it did, line
 * breaking would inherit whatever the browser decides, the server's renderer
 * would decide something slightly different, and the pages on screen would stop
 * matching the pages in the .docx — the one failure that makes a paginated
 * editor worse than an unpaginated one.
 *
 * Advance widths live in the font file. Reading them ourselves makes the
 * arithmetic deterministic and identical in a browser, in Node, and in a test,
 * and it is the same data any conforming renderer uses.
 *
 * ## Scope
 *
 * Only the tables text measurement needs. This is deliberately NOT a font
 * library: no glyph outlines, no rasterisation, no variable-font axes. Tables
 * we do not understand are left alone rather than guessed at.
 */

/** A four-character sfnt table tag, e.g. `head`. */
export type TableTag = string;

export interface TableRecord {
    readonly tag: TableTag;
    readonly offset: number;
    readonly length: number;
}

/**
 * Big-endian reader over the font bytes.
 *
 * Fonts are big-endian; `DataView` is the only sane way to read them in JS, and
 * every accessor here is bounds-checked because a truncated or hostile file
 * must fail as a clear error rather than as silently wrong metrics.
 */
export class FontReader {
    private readonly view: DataView;

    constructor(private readonly bytes: Uint8Array) {
        this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    }

    get byteLength(): number {
        return this.bytes.byteLength;
    }

    uint8(offset: number): number {
        this.require(offset, 1);

        return this.view.getUint8(offset);
    }

    uint16(offset: number): number {
        this.require(offset, 2);

        return this.view.getUint16(offset, false);
    }

    int16(offset: number): number {
        this.require(offset, 2);

        return this.view.getInt16(offset, false);
    }

    uint32(offset: number): number {
        this.require(offset, 4);

        return this.view.getUint32(offset, false);
    }

    /** Four bytes as ASCII — how sfnt spells its table names. */
    tag(offset: number): string {
        this.require(offset, 4);

        return String.fromCharCode(
            this.view.getUint8(offset),
            this.view.getUint8(offset + 1),
            this.view.getUint8(offset + 2),
            this.view.getUint8(offset + 3),
        );
    }

    private require(offset: number, length: number): void {
        if (offset < 0 || offset + length > this.bytes.byteLength) {
            throw new RangeError(
                `Font read out of bounds at ${offset}+${length} (file is ${this.bytes.byteLength} bytes). `
                + 'The file is truncated or is not a font.',
            );
        }
    }
}

/**
 * The sfnt table directory: what tables the file has and where they are.
 *
 * Accepts the two containers that matter: `0x00010000` (TrueType outlines) and
 * `OTTO` (CFF outlines). Both carry the same metric tables, which is all we
 * read — the outline format only decides how glyphs are DRAWN, and we never
 * draw them.
 *
 * TrueType Collections (`ttcf`) are refused rather than half-handled: a
 * collection holds several fonts and picking one silently would mean measuring
 * with a face nobody asked for.
 */
export function readTableDirectory(reader: FontReader): Map<TableTag, TableRecord> {
    const sfntVersion = reader.uint32(0);
    const TRUETYPE = 0x00010000;
    const OTTO = 0x4f54544f;
    const TTCF = 0x74746366;

    if (sfntVersion === TTCF) {
        throw new Error('TrueType Collections are not supported: a collection holds several faces and choosing one silently would measure with the wrong font.');
    }
    if (sfntVersion !== TRUETYPE && sfntVersion !== OTTO) {
        throw new Error(`Not an sfnt font: unexpected version 0x${sfntVersion.toString(16)}.`);
    }

    const numTables = reader.uint16(4);
    const tables = new Map<TableTag, TableRecord>();

    // The directory starts at 12; each record is 16 bytes:
    // tag(4) checksum(4) offset(4) length(4).
    for (let i = 0; i < numTables; i++) {
        const record = 12 + i * 16;
        const tag = reader.tag(record);
        const offset = reader.uint32(record + 8);
        const length = reader.uint32(record + 12);

        if (offset + length > reader.byteLength) {
            throw new Error(`Table "${tag}" claims bytes past the end of the file; the font is truncated.`);
        }

        tables.set(tag, { tag, offset, length });
    }

    return tables;
}

export function requireTable(tables: Map<TableTag, TableRecord>, tag: TableTag): TableRecord {
    const table = tables.get(tag);
    if (!table) {
        throw new Error(`Font is missing the "${tag}" table, which text measurement needs.`);
    }

    return table;
}
