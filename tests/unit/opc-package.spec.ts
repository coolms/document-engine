import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { describe, expect, it } from 'vitest';

import { OpcPackage, webStreamsCodec, type ZipCodec } from '../../src/opc/opc-package.js';
import { readZip, writeZip } from '../../src/opc/zip.js';
import { crc32 } from '../../src/opc/crc32.js';
import {
    findById,
    findByType,
    readRelationships,
    relationshipPartFor,
    resolveTarget,
} from '../../src/opc/relationships.js';

const FIXTURES = join(dirname(fileURLToPath(import.meta.url)), '../fixtures/docx');

function file(name: string): Uint8Array {
    return new Uint8Array(readFileSync(join(FIXTURES, name)));
}

const LEASE = 'lease-landscape.docx';
const AUTHORED = 'word-authored.docx';

interface HandBuiltZip {
    readonly name?: string;
    readonly content?: Uint8Array;
    readonly method?: number;
    readonly flags?: number;
    /** A local-header extra field, which the central directory need not match. */
    readonly localExtra?: Uint8Array;
    /** An archive comment, which pushes the end record away from the end. */
    readonly comment?: Uint8Array;
    /** Override the recorded size, to model a truncated archive. */
    readonly declaredSize?: number;
}

/**
 * Build a single-entry archive by hand.
 *
 * Real `.docx` files are well-behaved — no comments, no local extra fields, no
 * encryption — so they cannot exercise the format rules that exist for the
 * files that are not. Every option below is legal zip that a producer may emit
 * and that a reader has to survive.
 */
function handBuilt(options: HandBuiltZip = {}): Uint8Array {
    const name = new TextEncoder().encode(options.name ?? 'word/document.xml');
    const content = options.content ?? new TextEncoder().encode('<a/>');
    const extra = options.localExtra ?? new Uint8Array(0);
    const comment = options.comment ?? new Uint8Array(0);
    const method = options.method ?? 0;
    const flags = options.flags ?? 0;
    const size = options.declaredSize ?? content.length;

    const local = new Uint8Array(30 + name.length + extra.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(6, flags, true);
    localView.setUint16(8, method, true);
    localView.setUint32(14, crc32(content), true);
    localView.setUint32(18, size, true);
    localView.setUint32(22, content.length, true);
    localView.setUint16(26, name.length, true);
    localView.setUint16(28, extra.length, true);
    local.set(name, 30);
    local.set(extra, 30 + name.length);

    const central = new Uint8Array(46 + name.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(8, flags, true);
    centralView.setUint16(10, method, true);
    centralView.setUint32(16, crc32(content), true);
    centralView.setUint32(20, size, true);
    centralView.setUint32(24, content.length, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint32(42, 0, true);
    central.set(name, 46);

    const end = new Uint8Array(22 + comment.length);
    const endView = new DataView(end.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(8, 1, true);
    endView.setUint16(10, 1, true);
    endView.setUint32(12, central.length, true);
    endView.setUint32(16, local.length + content.length, true);
    endView.setUint16(20, comment.length, true);
    end.set(comment, 22);

    const out = new Uint8Array(local.length + content.length + central.length + end.length);
    out.set(local, 0);
    out.set(content, local.length);
    out.set(central, local.length + content.length);
    out.set(end, local.length + content.length + central.length);

    return out;
}

describe('crc32', () => {
    it('agrees with the standard check value', () => {
        // 0xCBF43926 for "123456789" is the published CRC-32 check value, so
        // this disagrees with the implementation rather than restating it.
        expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926);
        expect(crc32(new Uint8Array(0))).toBe(0);
    });

    it('stays unsigned, as a zip header requires', () => {
        // Without the final unsigned shift most inputs come out negative and
        // the archive is unopenable.
        for (const text of ['a', 'hello world', 'Привет', 'x'.repeat(1000)]) {
            expect(crc32(new TextEncoder().encode(text))).toBeGreaterThanOrEqual(0);
        }
    });
});

describe('relationships', () => {
    it('resolves a target against the folder of the part that declared it', () => {
        // "styles.xml" declared by word/document.xml is word/styles.xml, not
        // styles.xml at the package root.
        expect(resolveTarget('word/document.xml', 'styles.xml')).toBe('word/styles.xml');
        expect(resolveTarget('', 'word/document.xml')).toBe('word/document.xml');
        expect(resolveTarget('word/document.xml', '/word/other.xml')).toBe('word/other.xml');
        expect(resolveTarget('word/embeddings/x.xml', '../media/image1.png')).toBe('word/media/image1.png');
        expect(resolveTarget('word/document.xml', './styles.xml')).toBe('word/styles.xml');
    });

    it('names the relationship part for a part', () => {
        expect(relationshipPartFor('word/document.xml')).toBe('word/_rels/document.xml.rels');
        expect(relationshipPartFor('')).toBe('_rels/.rels');
    });

    it('finds a relationship by Id, and never an EXTERNAL one', () => {
        // A header reference names an r:id. An external target is a URL rather
        // than a part, and resolving one as a part looks for a file that is not
        // in the package.
        const relationships = readRelationships(
            '<Relationships xmlns="x">'
            + '<Relationship Id="rId7" Type=".../header" Target="header1.xml"/>'
            + '<Relationship Id="rId8" Type=".../hyperlink" Target="http://e.com" TargetMode="External"/>'
            + '</Relationships>',
        );

        expect(findById(relationships, 'rId7')?.target).toBe('header1.xml');
        expect(findById(relationships, 'rId8')).toBeNull();
        expect(findById(relationships, 'rId9')).toBeNull();
    });

    it('matches a relationship type by its LAST segment', () => {
        // The strict and transitional flavours of the format use different
        // namespace URIs for the same relationship; a whole-string match finds
        // neither on a strict document.
        const relationships = readRelationships(
            '<Relationships xmlns="x">'
            + '<Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"'
            + ' Target="word/document.xml"/>'
            + '<Relationship Id="rId2" Type="http://.../relationships/hyperlink" Target="http://e.com"'
            + ' TargetMode="External"/>'
            + '</Relationships>',
        );

        expect(findByType(relationships, 'officeDocument')?.target).toBe('word/document.xml');
        // An external target is a URL, not a part, and must never be resolved
        // as one.
        expect(findByType(relationships, 'hyperlink')).toBeNull();
    });
});

describe('OpcPackage', () => {
    describe('reading', () => {
        it('lists the parts of a real .docx', async () => {
            const opc = OpcPackage.open(file(AUTHORED));

            expect(opc.names).toContain('word/document.xml');
            expect(opc.names).toContain('[Content_Types].xml');
            expect(opc.names).toContain('_rels/.rels');
            // Directory entries carry no information an OPC package needs; the
            // part names hold the paths.
            expect(opc.names.filter((name) => name.endsWith('/'))).toEqual([]);
        });

        it('inflates a part to exactly what the archive says it should be', async () => {
            const opc = OpcPackage.open(file(AUTHORED));
            const xml = await opc.text('word/document.xml');

            expect(xml.length).toBe(48484);
            expect(xml.startsWith('<?xml')).toBe(true);
            expect(xml).toContain('<w:document');
        });

        it('decompresses NOTHING until a part is asked for', async () => {
            // A document is opened far more often than every part of it is
            // read, and inflating a whole package to answer "what parts are
            // there" is work nobody asked for.
            let inflations = 0;
            const counting: ZipCodec = {
                inflateRaw: (bytes) => {
                    inflations++;

                    return webStreamsCodec.inflateRaw(bytes);
                },
                deflateRaw: (bytes) => webStreamsCodec.deflateRaw(bytes),
            };

            const opc = OpcPackage.open(file(AUTHORED), counting);
            expect(opc.names.length).toBeGreaterThan(10);
            expect(inflations).toBe(0);

            await opc.text('word/document.xml');
            expect(inflations).toBe(1);

            // And each part only once, however often it is read.
            await opc.text('word/document.xml');
            expect(inflations).toBe(1);
        });

        it('reports a missing part by name, and says what it does have', async () => {
            const opc = OpcPackage.open(file(AUTHORED));

            await expect(opc.text('word/nonexistent.xml')).rejects.toThrow(/no part "word\/nonexistent.xml"/);
            expect(await opc.textIfPresent('word/nonexistent.xml')).toBeNull();
        });
    });

    describe('saving', () => {
        it('reproduces every part when nothing was changed', async () => {
            const original = file(AUTHORED);
            const opc = OpcPackage.open(original);
            const saved = await opc.save();

            const before = new Map(readZip(original).map((entry) => [entry.name, entry]));
            const after = new Map(readZip(saved).map((entry) => [entry.name, entry]));

            expect([...after.keys()]).toEqual([...before.keys()]);
            for (const [name, entry] of before) {
                expect(after.get(name)!.crc).toBe(entry.crc);
                expect(after.get(name)!.uncompressedSize).toBe(entry.uncompressedSize);
            }
        });

        it('keeps the ORIGINAL COMPRESSED BYTES of every untouched part', async () => {
            // Not merely equivalent content — the same bytes. Re-compressing
            // would change every part of the file even where nothing changed,
            // which throws away exactly what the XML layer's byte-exact surgery
            // is for.
            const original = file(AUTHORED);
            const opc = OpcPackage.open(original);
            opc.replace('word/document.xml', await opc.text('word/document.xml') + '<!-- x -->');

            const saved = await opc.save();
            const before = new Map(readZip(original).map((entry) => [entry.name, entry]));

            for (const entry of readZip(saved)) {
                if ('word/document.xml' === entry.name) {
                    continue;
                }
                expect([...entry.compressed]).toEqual([...before.get(entry.name)!.compressed]);
            }
        });

        it('is byte-identical when saved twice', async () => {
            // No timestamp of "now" anywhere: saving must be deterministic, or
            // every save shows up as a whole-file change.
            const opc = OpcPackage.open(file(LEASE));
            const first = await opc.save();
            const second = await opc.save();

            expect([...second]).toEqual([...first]);
        });

        it('round-trips a replaced part through a real save and re-open', async () => {
            const opc = OpcPackage.open(file(LEASE));
            const edited = (await opc.text('word/document.xml')).replace('Договор аренды', 'ЗАМЕНА');
            opc.replace('word/document.xml', edited);

            const reopened = OpcPackage.open(await opc.save());

            expect(await reopened.text('word/document.xml')).toBe(edited);
            expect(await reopened.text('word/styles.xml')).toBe(await opc.text('word/styles.xml'));
        });

        it('accepts bytes as well as text', async () => {
            const opc = OpcPackage.open(file(LEASE));
            const content = new TextEncoder().encode('<a/>');
            opc.replace('word/document.xml', content);

            expect(await OpcPackage.open(await opc.save()).text('word/document.xml')).toBe('<a/>');
        });

        it('stores content that deflate would make LARGER', async () => {
            // Very short content compresses to more than it started as. Storing
            // it keeps the package smaller than the thing it describes.
            const opc = OpcPackage.open(file(LEASE));
            opc.replace('word/document.xml', 'x');

            const entry = readZip(await opc.save()).find((e) => 'word/document.xml' === e.name)!;
            expect(entry.method).toBe(0);
            expect(entry.compressed.length).toBe(1);
        });

        it('keeps a replaced part\'s original timestamp, so a save is reproducible', async () => {
            // Stamping "now" would make every save differ from the last, and a
            // document's real modification time lives in docProps/core.xml
            // rather than in the container.
            const original = file(LEASE);
            const before = readZip(original).find((e) => 'word/document.xml' === e.name)!;

            const opc = OpcPackage.open(original);
            opc.replace('word/document.xml', 'x');
            const after = readZip(await opc.save()).find((e) => 'word/document.xml' === e.name)!;

            expect(after.dosDate).toBe(before.dosDate);
            expect(after.dosTime).toBe(before.dosTime);
        });

        it('reads back a replacement without saving first', async () => {
            // The editor changes a part and then reads it again — through the
            // same open package. Returning the file's original content there
            // would silently discard the edit.
            const opc = OpcPackage.open(file(LEASE));
            opc.replace('word/document.xml', '<replaced/>');

            expect(await opc.text('word/document.xml')).toBe('<replaced/>');
        });

        it('says which parts it has been told to change', async () => {
            const opc = OpcPackage.open(file(LEASE));

            expect(opc.isModified('word/document.xml')).toBe(false);
            opc.replace('word/document.xml', 'x');
            expect(opc.isModified('word/document.xml')).toBe(true);
            expect(opc.isModified('word/styles.xml')).toBe(false);
        });

        it('refuses to add a part it cannot declare a content type for', async () => {
            // A part added without a [Content_Types].xml entry is invisible to
            // Word — a silent no-op that looks like it worked.
            const opc = OpcPackage.open(file(LEASE));

            expect(() => opc.replace('word/brand-new.xml', '<a/>')).toThrow(/has no such part/);
        });
    });

    describe('archives that are legal but not tidy', () => {
        it('finds the data past a local EXTRA FIELD the central directory lacks', async () => {
            // The two headers are allowed to carry different extra fields, and
            // several producers make use of that. Trusting the central
            // directory's lengths starts the read a few bytes into the file
            // data, which inflates to rubbish rather than failing.
            const archive = handBuilt({ localExtra: new Uint8Array([0x55, 0x54, 4, 0, 1, 2, 3, 4]) });

            expect(await OpcPackage.open(archive).text('word/document.xml')).toBe('<a/>');
        });

        it('finds the end record behind an archive COMMENT', async () => {
            // The end-of-central-directory record is last only when there is no
            // comment; a comment of up to 65535 bytes may follow it.
            const archive = handBuilt({ comment: new TextEncoder().encode('written by something else') });

            expect(await OpcPackage.open(archive).text('word/document.xml')).toBe('<a/>');
        });

        it('refuses an encrypted entry instead of inflating the ciphertext', () => {
            const archive = handBuilt({ flags: 0x0001 });

            expect(() => OpcPackage.open(archive)).toThrow(/encrypted/);
        });

        it('refuses a compression method it cannot read', () => {
            // Naming the method matters: "cannot open" is not a diagnosis.
            const archive = handBuilt({ method: 12 });

            expect(() => OpcPackage.open(archive)).toThrow(/compression method 12/);
        });

        it('refuses an entry whose data runs past the end of the file', () => {
            const archive = handBuilt({ declaredSize: 9999 });

            expect(() => OpcPackage.open(archive)).toThrow(/runs past the end|truncated/);
        });
    });

    describe('damaged input', () => {
        it('refuses bytes that are not a zip', () => {
            expect(() => OpcPackage.open(new Uint8Array(100)))
                .toThrow(/no end-of-central-directory/);
        });

        it('refuses a truncated archive rather than reading past the end', () => {
            const truncated = file(LEASE).slice(0, 2000);

            expect(() => OpcPackage.open(truncated)).toThrow();
        });

        it('refuses a part whose content does not match its recorded checksum', async () => {
            // Built deliberately, because corrupting a DEFLATE stream usually
            // fails inflation before the checksum is ever consulted. A stored
            // entry has no inflation to fail, so this reaches the check itself.
            //
            // Silent corruption is worse than a refusal: the damage propagates
            // into everything saved afterwards.
            const content = new TextEncoder().encode('<a/>');
            const archive = writeZip([{
                name: 'word/document.xml',
                method: 0,
                crc: crc32(content) ^ 0xffff, // deliberately wrong
                uncompressedSize: content.length,
                compressed: content,
                dosTime: 0,
                dosDate: 0x21,
                externalAttributes: 0,
            }]);

            await expect(OpcPackage.open(archive).text('word/document.xml'))
                .rejects.toThrow(/fails its checksum/);
        });

        it('refuses a part whose recorded LENGTH is wrong', async () => {
            const content = new TextEncoder().encode('<a/>');
            const archive = writeZip([{
                name: 'word/document.xml',
                method: 0,
                crc: crc32(content),
                uncompressedSize: 999,
                compressed: content,
                dosTime: 0,
                dosDate: 0x21,
                externalAttributes: 0,
            }]);

            await expect(OpcPackage.open(archive).text('word/document.xml'))
                .rejects.toThrow(/inflated to 4 bytes, not the 999 recorded/);
        });

        it('fails on a damaged compressed stream rather than returning rubbish', async () => {
            const bytes = file(LEASE);
            const entry = readZip(bytes).find((e) => 'word/document.xml' === e.name)!;

            // subarray keeps the offset into the original buffer, so this hits
            // the part's own bytes rather than some byte that happens to match.
            const damaged = bytes.slice();
            const at = entry.compressed.byteOffset + 10;
            damaged[at] = damaged[at]! ^ 0xff;

            await expect(OpcPackage.open(damaged).text('word/document.xml')).rejects.toThrow();
        });
    });
});
