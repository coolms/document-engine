import { crc32 } from './crc32.js';
import {
    METHOD_DEFLATE,
    METHOD_STORED,
    concat,
    readZip,
    verify,
    writeZip,
    type ZipEntry,
    type ZipEntryToWrite,
} from './zip.js';

/**
 * An OOXML package — a `.docx`, `.xlsx` or `.pptx` — as a set of named parts.
 *
 * This is the layer that makes the engine usable on a FILE rather than on a
 * part someone already extracted. Together with the XML layer it gives the
 * whole chain the same property end to end: open a document, change one
 * paragraph, save it, and every byte you did not touch is the byte the original
 * producer wrote — inside the parts, because the XML layer preserves them, and
 * between the parts, because an untouched entry is re-emitted with its original
 * COMPRESSED bytes rather than compressed again.
 */

export interface ZipCodec {
    inflateRaw(bytes: Uint8Array): Promise<Uint8Array>;
    deflateRaw(bytes: Uint8Array): Promise<Uint8Array>;
}

/**
 * The one codec, working in Node and in a browser unchanged.
 *
 * `DecompressionStream` is a platform API in both, so the package needs no
 * dependency and no `node:zlib` import — an import that would work here and
 * break the moment this is bundled for the editor.
 */
export const webStreamsCodec: ZipCodec = {
    inflateRaw: (bytes) => pipe(bytes, new DecompressionStream('deflate-raw')),
    deflateRaw: (bytes) => pipe(bytes, new CompressionStream('deflate-raw')),
};

async function pipe(bytes: Uint8Array, transform: DecompressionStream | CompressionStream): Promise<Uint8Array> {
    const source = new ReadableStream<Uint8Array>({
        start(controller): void {
            controller.enqueue(bytes);
            controller.close();
        },
    });

    const reader = source
        .pipeThrough(transform as unknown as ReadableWritablePair<Uint8Array, Uint8Array>)
        .getReader();

    const chunks: Uint8Array[] = [];
    for (;;) {
        const { done, value } = await reader.read();
        if (done) {
            break;
        }
        chunks.push(value);
    }

    return concat(chunks);
}

export class OpcPackage {
    private readonly entries = new Map<string, ZipEntry>();
    /** Part order as the file stored it, so saving does not reshuffle it. */
    private readonly order: string[] = [];
    private readonly inflated = new Map<string, Uint8Array>();
    private readonly replacements = new Map<string, Uint8Array>();

    private constructor(private readonly codec: ZipCodec) {}

    /**
     * Parse the container. Synchronous: nothing is decompressed until a part is
     * actually read, and a document is opened far more often than every part of
     * it is needed.
     */
    static open(bytes: Uint8Array, codec: ZipCodec = webStreamsCodec): OpcPackage {
        const opc = new OpcPackage(codec);
        for (const entry of readZip(bytes)) {
            opc.entries.set(entry.name, entry);
            opc.order.push(entry.name);
        }

        return opc;
    }

    get names(): readonly string[] {
        return this.order;
    }

    has(name: string): boolean {
        return this.entries.has(name);
    }

    async bytes(name: string): Promise<Uint8Array> {
        const replaced = this.replacements.get(name);
        if (undefined !== replaced) {
            return replaced;
        }

        const cached = this.inflated.get(name);
        if (undefined !== cached) {
            return cached;
        }

        const entry = this.entries.get(name);
        if (undefined === entry) {
            throw new Error(`Package has no part "${name}". It holds: ${this.order.join(', ')}`);
        }

        const content = METHOD_STORED === entry.method
            ? entry.compressed.slice()
            : await this.codec.inflateRaw(entry.compressed);

        verify(entry, content);
        this.inflated.set(name, content);

        return content;
    }

    async text(name: string): Promise<string> {
        return new TextDecoder('utf-8').decode(await this.bytes(name));
    }

    /** The part if it is there, or null — for the many OPC parts that are optional. */
    async textIfPresent(name: string): Promise<string | null> {
        return this.has(name) ? this.text(name) : null;
    }

    /**
     * Replace a part's content.
     *
     * Only an EXISTING part: adding one means declaring its content type in
     * `[Content_Types].xml` and usually a relationship as well, and a part
     * added without those is invisible to Word — a silent no-op that looks like
     * it worked.
     */
    replace(name: string, content: Uint8Array | string): void {
        if (!this.entries.has(name)) {
            throw new Error(
                `Cannot replace "${name}": the package has no such part, and adding one needs a `
                + `[Content_Types].xml declaration this layer does not manage.`,
            );
        }

        this.replacements.set(
            name,
            'string' === typeof content ? new TextEncoder().encode(content) : content,
        );
    }

    isModified(name: string): boolean {
        return this.replacements.has(name);
    }

    /**
     * Write the package back.
     *
     * Untouched parts keep their original compressed bytes, checksum, timestamp
     * and order; only what was replaced is compressed again. Saving a package
     * nobody edited therefore reproduces the input.
     */
    async save(): Promise<Uint8Array> {
        const out: ZipEntryToWrite[] = [];

        for (const name of this.order) {
            const entry = this.entries.get(name)!;
            const replacement = this.replacements.get(name);

            if (undefined === replacement) {
                out.push(entry);
                continue;
            }

            const deflated = await this.codec.deflateRaw(replacement);
            // Deflate can make very short content LONGER. Storing it then is
            // what a real zip writer does, and it keeps the package smaller
            // than the thing it describes.
            const useDeflate = deflated.length < replacement.length;

            out.push({
                name,
                method: useDeflate ? METHOD_DEFLATE : METHOD_STORED,
                crc: crc32(replacement),
                uncompressedSize: replacement.length,
                compressed: useDeflate ? deflated : replacement,
                // The original timestamp is kept rather than "now": saving twice
                // should produce the same bytes, and a document's real
                // modification time lives in docProps/core.xml anyway.
                dosTime: entry.dosTime,
                dosDate: entry.dosDate,
                externalAttributes: entry.externalAttributes,
            });
        }

        return writeZip(out);
    }
}
