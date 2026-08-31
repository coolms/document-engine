/**
 * CRC-32, the checksum a zip entry carries.
 *
 * Needed only for parts this engine REWRITES: an untouched entry keeps the
 * checksum the original producer wrote, alongside its original compressed
 * bytes. A wrong CRC is not a soft failure — Word and every zip tool reject the
 * whole package, so this is verified against the standard check value rather
 * than against itself.
 */

const POLYNOMIAL = 0xedb88320;

const TABLE = ((): Uint32Array => {
    const table = new Uint32Array(256);
    for (let index = 0; index < 256; index++) {
        let value = index;
        for (let bit = 0; bit < 8; bit++) {
            value = 0 !== (value & 1) ? POLYNOMIAL ^ (value >>> 1) : value >>> 1;
        }
        table[index] = value >>> 0;
    }

    return table;
})();

export function crc32(bytes: Uint8Array): number {
    let crc = 0xffffffff;
    for (let index = 0; index < bytes.length; index++) {
        crc = TABLE[(crc ^ bytes[index]!) & 0xff]! ^ (crc >>> 8);
    }

    // The unsigned shift matters: without it this is negative for most inputs
    // and writing it into a zip header produces a file nothing will open.
    return (crc ^ 0xffffffff) >>> 0;
}
