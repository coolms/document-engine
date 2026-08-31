import { crc32 } from './crc32.js';

/**
 * The zip container an OOXML package is.
 *
 * ## Why the compressed bytes are kept
 *
 * An entry nobody edited is re-emitted with the EXACT compressed bytes it
 * arrived with — never re-compressed. Two reasons, and the second is the
 * important one:
 *
 * 1. Re-compressing is work nobody asked for on a package where one part
 *    changed and forty did not.
 * 2. Deflate output depends on the implementation and its level, so
 *    re-compressing changes every byte of every part even when the content is
 *    identical. That would defeat the whole point of the XML layer's byte-exact
 *    surgery: it preserves the bytes, and this would scramble them again on the
 *    way out.
 *
 * ## What is deliberately not supported
 *
 * Encryption, multi-disk archives, and compression methods other than stored
 * and deflate — none of which appear in an OOXML package. Each is REFUSED with
 * a clear message rather than mis-parsed into plausible rubbish.
 */

const LOCAL_HEADER = 0x04034b50;
const CENTRAL_HEADER = 0x02014b50;
const END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const ZIP64_END_OF_CENTRAL_DIRECTORY = 0x06064b50;
const ZIP64_LOCATOR = 0x07064b50;

export const METHOD_STORED = 0;
export const METHOD_DEFLATE = 8;

export interface ZipEntry {
    readonly name: string;
    readonly method: number;
    readonly crc: number;
    readonly uncompressedSize: number;
    /** Exactly as stored in the source archive. */
    readonly compressed: Uint8Array;
    /** MS-DOS date and time, preserved so an untouched package round-trips. */
    readonly dosTime: number;
    readonly dosDate: number;
    /** External attributes — carries the unix mode some producers set. */
    readonly externalAttributes: number;
}

export function readZip(bytes: Uint8Array): ZipEntry[] {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const directory = findCentralDirectory(bytes, view);

    const entries: ZipEntry[] = [];
    let offset = directory.offset;

    for (let index = 0; index < directory.count; index++) {
        if (CENTRAL_HEADER !== view.getUint32(offset, true)) {
            throw new Error(`Zip central directory entry ${index} has a bad signature; the archive is damaged.`);
        }

        const flags = view.getUint16(offset + 8, true);
        if (0 !== (flags & 0x0001)) {
            throw new Error(`Zip entry is encrypted; an OOXML package this engine can open never is.`);
        }

        const method = view.getUint16(offset + 10, true);
        const dosTime = view.getUint16(offset + 12, true);
        const dosDate = view.getUint16(offset + 14, true);
        const crc = view.getUint32(offset + 16, true);
        const compressedSize = view.getUint32(offset + 20, true);
        const uncompressedSize = view.getUint32(offset + 24, true);
        const nameLength = view.getUint16(offset + 28, true);
        const extraLength = view.getUint16(offset + 30, true);
        const commentLength = view.getUint16(offset + 32, true);
        const externalAttributes = view.getUint32(offset + 38, true);
        const localOffset = view.getUint32(offset + 42, true);

        const name = decodeName(bytes.subarray(offset + 46, offset + 46 + nameLength));

        if (METHOD_STORED !== method && METHOD_DEFLATE !== method) {
            throw new Error(`Zip entry "${name}" uses compression method ${method}; only stored and deflate are read.`);
        }
        if (0xffffffff === compressedSize || 0xffffffff === uncompressedSize || 0xffffffff === localOffset) {
            throw new Error(`Zip entry "${name}" needs ZIP64 field sizes, which this reader does not handle.`);
        }

        // Directories are recorded as entries with no content. They carry no
        // information an OPC package needs — the part names hold the paths.
        if (!name.endsWith('/')) {
            entries.push({
                name,
                method,
                crc,
                uncompressedSize,
                compressed: dataOf(bytes, view, localOffset, compressedSize, name),
                dosTime,
                dosDate,
                externalAttributes,
            });
        }

        offset += 46 + nameLength + extraLength + commentLength;
    }

    return entries;
}

/**
 * Locate the entry's data.
 *
 * The local header's name and extra-field lengths are read rather than reused
 * from the central directory: they are ALLOWED to differ, and several producers
 * write a different extra field in each. Assuming they match puts the read a
 * few bytes into the file data, which inflates to garbage rather than failing.
 */
function dataOf(
    bytes: Uint8Array,
    view: DataView,
    localOffset: number,
    compressedSize: number,
    name: string,
): Uint8Array {
    if (LOCAL_HEADER !== view.getUint32(localOffset, true)) {
        throw new Error(`Zip entry "${name}" has no local header at its recorded offset; the archive is damaged.`);
    }

    const start = localOffset + 30
        + view.getUint16(localOffset + 26, true)
        + view.getUint16(localOffset + 28, true);

    if (start + compressedSize > bytes.length) {
        throw new Error(`Zip entry "${name}" runs past the end of the archive; the file is truncated.`);
    }

    return bytes.subarray(start, start + compressedSize);
}

function findCentralDirectory(bytes: Uint8Array, view: DataView): { offset: number; count: number } {
    // The end record is last, but a trailing comment of up to 65535 bytes may
    // follow it, so it has to be searched for rather than assumed.
    const earliest = Math.max(0, bytes.length - 0xffff - 22);
    let end = -1;
    for (let index = bytes.length - 22; index >= earliest; index--) {
        if (END_OF_CENTRAL_DIRECTORY === view.getUint32(index, true)) {
            end = index;
            break;
        }
    }

    if (end < 0) {
        throw new Error('Not a zip archive: no end-of-central-directory record.');
    }

    let count = view.getUint16(end + 10, true);
    let offset = view.getUint32(end + 16, true);

    if (0xffff === count || 0xffffffff === offset) {
        ({ count, offset } = readZip64(view, end));
    }

    return { offset, count };
}

/**
 * ZIP64, which some producers emit even for small archives.
 *
 * Reading it is a few lines; refusing it would reject packages that open
 * everywhere else.
 */
function readZip64(view: DataView, endOffset: number): { offset: number; count: number } {
    const locator = endOffset - 20;
    if (locator < 0 || ZIP64_LOCATOR !== view.getUint32(locator, true)) {
        throw new Error('Zip archive claims ZIP64 sizes but carries no ZIP64 locator.');
    }

    const record = Number(view.getBigUint64(locator + 8, true));
    if (ZIP64_END_OF_CENTRAL_DIRECTORY !== view.getUint32(record, true)) {
        throw new Error('Zip ZIP64 locator does not point at a ZIP64 end-of-central-directory record.');
    }

    return {
        count: Number(view.getBigUint64(record + 32, true)),
        offset: Number(view.getBigUint64(record + 48, true)),
    };
}

export interface ZipEntryToWrite {
    readonly name: string;
    readonly method: number;
    readonly crc: number;
    readonly uncompressedSize: number;
    readonly compressed: Uint8Array;
    readonly dosTime: number;
    readonly dosDate: number;
    readonly externalAttributes: number;
}

export function writeZip(entries: readonly ZipEntryToWrite[]): Uint8Array {
    const chunks: Uint8Array[] = [];
    const central: Uint8Array[] = [];
    let offset = 0;

    for (const entry of entries) {
        const name = encodeName(entry.name);

        const local = new Uint8Array(30 + name.length);
        const localView = new DataView(local.buffer);
        localView.setUint32(0, LOCAL_HEADER, true);
        localView.setUint16(4, 20, true); // version needed: 2.0, deflate
        localView.setUint16(6, 0x0800, true); // UTF-8 names
        localView.setUint16(8, entry.method, true);
        localView.setUint16(10, entry.dosTime, true);
        localView.setUint16(12, entry.dosDate, true);
        localView.setUint32(14, entry.crc, true);
        localView.setUint32(18, entry.compressed.length, true);
        localView.setUint32(22, entry.uncompressedSize, true);
        localView.setUint16(26, name.length, true);
        localView.setUint16(28, 0, true);
        local.set(name, 30);

        const header = new Uint8Array(46 + name.length);
        const headerView = new DataView(header.buffer);
        headerView.setUint32(0, CENTRAL_HEADER, true);
        headerView.setUint16(4, 20, true);
        headerView.setUint16(6, 20, true);
        headerView.setUint16(8, 0x0800, true);
        headerView.setUint16(10, entry.method, true);
        headerView.setUint16(12, entry.dosTime, true);
        headerView.setUint16(14, entry.dosDate, true);
        headerView.setUint32(16, entry.crc, true);
        headerView.setUint32(20, entry.compressed.length, true);
        headerView.setUint32(24, entry.uncompressedSize, true);
        headerView.setUint16(28, name.length, true);
        headerView.setUint32(42, offset, true);
        headerView.setUint32(38, entry.externalAttributes, true);
        header.set(name, 46);

        chunks.push(local, entry.compressed);
        central.push(header);
        offset += local.length + entry.compressed.length;
    }

    const centralSize = central.reduce((sum, part) => sum + part.length, 0);
    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    endView.setUint32(0, END_OF_CENTRAL_DIRECTORY, true);
    endView.setUint16(8, entries.length, true);
    endView.setUint16(10, entries.length, true);
    endView.setUint32(12, centralSize, true);
    endView.setUint32(16, offset, true);

    return concat([...chunks, ...central, end]);
}

export function concat(parts: readonly Uint8Array[]): Uint8Array {
    const total = parts.reduce((sum, part) => sum + part.length, 0);
    const out = new Uint8Array(total);
    let at = 0;
    for (const part of parts) {
        out.set(part, at);
        at += part.length;
    }

    return out;
}

/**
 * Verify an entry's content against the checksum the archive recorded.
 *
 * Silent corruption in a document is worse than a refusal to open it: the
 * damage propagates into everything saved afterwards.
 */
export function verify(entry: ZipEntry, content: Uint8Array): void {
    if (content.length !== entry.uncompressedSize) {
        throw new Error(
            `Zip entry "${entry.name}" inflated to ${content.length} bytes, not the ${entry.uncompressedSize} recorded.`,
        );
    }
    const actual = crc32(content);
    if (actual !== entry.crc) {
        throw new Error(`Zip entry "${entry.name}" fails its checksum; the archive is corrupt.`);
    }
}

const decoder = new TextDecoder('utf-8');
const encoder = new TextEncoder();

function decodeName(bytes: Uint8Array): string {
    return decoder.decode(bytes);
}

function encodeName(name: string): Uint8Array {
    return encoder.encode(name);
}
