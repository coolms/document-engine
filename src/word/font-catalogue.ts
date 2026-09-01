import { TrueTypeFont } from '../font/truetype-font.js';

/**
 * Turning the font a document ASKS for into a font we actually have.
 *
 * A `.docx` names Times New Roman, Calibri or Arial — faces we cannot ship and
 * cannot serve to a browser. The vendored set is metric-compatible with each of
 * them, which is the whole point: substituting Carlito for Calibri keeps every
 * advance width identical, so an inherited document does not repaginate when it
 * is opened here.
 *
 * The manifest is the single source of truth for that mapping and is read by
 * the PHP render side as well. Hard-coding the same table in two languages is
 * how the browser and the renderer end up disagreeing about where a page ends.
 */

export interface FontFileEntry {
    /**
     * The key this face is READ by, unique across every family in the manifest.
     *
     * A filename for the vendored set, which is what the directory holds. For a
     * face installed at runtime it is the row's id: two families can both
     * arrive as `Regular.ttf`, and this is what a byte cache and a
     * {@link FontFileReader} are keyed by.
     */
    readonly file: string;
    readonly sha256: string;
    readonly bytes: number;
    /**
     * Where the bytes come from, when they are not a file beside the manifest.
     *
     * Absent for everything the platform vendors -- those are static assets and
     * stay on the static path. Present for a face served by an endpoint, which
     * is how a font installed after the bundle was built reaches the canvas.
     */
    readonly url?: string;
    /** The name a person would recognise, for a list or a network panel. */
    readonly name?: string;
}

export interface FontFamilyEntry {
    /** Names in a document that should resolve to this family. */
    readonly substitutes: readonly string[];
    readonly files: Readonly<Record<string, FontFileEntry | undefined>>;
}

export interface FontManifest {
    readonly families: Readonly<Record<string, FontFamilyEntry | undefined>>;
    readonly defaults: { readonly family: string; readonly sizePt: number };
}

export type FontStyleKey = 'regular' | 'bold' | 'italic' | 'boldItalic';

export function fontStyleKey(bold: boolean, italic: boolean): FontStyleKey {
    if (bold && italic) {
        return 'boldItalic';
    }

    return bold ? 'bold' : italic ? 'italic' : 'regular';
}

export interface ResolvedFont {
    readonly font: TrueTypeFont;
    /** The family actually used. */
    readonly family: string;
    /** What the document asked for, when that was something else. */
    readonly substitutedFor: string | null;
}

/** Reads a font file's bytes. Injected so this works in a browser too. */
export type FontFileReader = (file: string) => Uint8Array;

export class FontCatalogue {
    /** Lower-cased requested name to family name. */
    private readonly aliases = new Map<string, string>();
    private readonly parsed = new Map<string, TrueTypeFont>();

    private constructor(
        private readonly manifest: FontManifest,
        private readonly read: FontFileReader,
    ) {
        for (const [family, entry] of Object.entries(manifest.families)) {
            if (undefined === entry) {
                continue;
            }
            this.aliases.set(family.toLowerCase(), family);
            for (const alias of entry.substitutes) {
                this.aliases.set(alias.toLowerCase(), family);
            }
        }
    }

    static load(manifest: FontManifest, read: FontFileReader): FontCatalogue {
        return new FontCatalogue(manifest, read);
    }

    get defaultFamily(): string {
        return this.manifest.defaults.family;
    }

    get defaultSizePt(): number {
        return this.manifest.defaults.sizePt;
    }

    /** True when a document could name this and get a real match. */
    knows(family: string): boolean {
        return this.aliases.has(family.trim().toLowerCase());
    }

    /**
     * Resolve a requested family and weight to a font we hold.
     *
     * An unknown family falls back to the default rather than throwing: a
     * document naming some font nobody has must still lay out. The caller is
     * told, through `substitutedFor`, so the substitution can be reported
     * instead of silently changing the page count.
     */
    resolve(requested: string | null, bold: boolean, italic: boolean): ResolvedFont {
        const asked = requested?.trim() ?? '';
        const family = this.aliases.get(asked.toLowerCase()) ?? this.defaultFamily;
        const entry = this.manifest.families[family];
        if (undefined === entry) {
            throw new Error(`Font manifest has no family "${family}"; it cannot be the default either.`);
        }

        const wanted = fontStyleKey(bold, italic);
        // Falling back to regular is a last resort and a metric lie — a regular
        // face is narrower than its bold. Every vendored family carries all
        // four, so this only fires if the manifest is incomplete.
        const file = entry.files[wanted] ?? entry.files['regular'];
        if (undefined === file) {
            throw new Error(`Font family "${family}" has no ${wanted} and no regular file.`);
        }

        return {
            font: this.parse(file.file),
            family,
            substitutedFor: '' !== asked && family.toLowerCase() !== asked.toLowerCase() ? asked : null,
        };
    }

    /** Parsed once per file — a document repeats the same few faces constantly. */
    private parse(file: string): TrueTypeFont {
        const cached = this.parsed.get(file);
        if (undefined !== cached) {
            return cached;
        }

        const font = TrueTypeFont.parse(this.read(file));
        this.parsed.set(file, font);

        return font;
    }
}

/**
 * Where the app's asset pipeline publishes this manifest for the browser.
 *
 *  An app-layout detail in an engine that never fetches anything, and it is
 * here on purpose: TWO browser packages read the manifest to build their font
 * select, and a path spelled out in each is a string that drifts the day the
 * pipeline changes. The engine owns the manifest, so it owns its published
 * name. Nothing in this package reads it -- see `theme-admin/angular`'s
 * `prebuild`, which is what actually puts the file there.
 */
export const FONT_MANIFEST_ASSET = 'assets/document-fonts/fonts.manifest.json';

/**
 * The families an editor should OFFER, derived from the manifest.
 *
 * ## Why this is derived rather than typed out
 *
 * A toolbar that states its own list drifts from what the platform can
 * actually measure, and the drift is silent: the editor offers Georgia, the
 * canvas paints whatever the author's machine has, the engine measures the
 * base face because nothing resolved, and LibreOffice prints a third font.
 * That was true for Georgia, Verdana and Tahoma for as long as the list was a
 * literal (one change named it, a later one removed the literal).
 *
 * So: installing a family is adding files and a manifest entry. Nothing in
 * either editor knows the names.
 *
 * The offered name is the FIRST substitute, because that is the name a
 * document should write -- `Calibri`, not `Carlito`. Word and LibreOffice both
 * resolve it, and `TextMapper::inlineStyle()` carries it into the .docx
 * unchanged. The vendored name is a fallback for an entry that lists no
 * substitutes at all, which the shipped manifest never does.
 *
 * The default family leads. It is the paper's own face and the one the
 * composer states, so it is the useful first choice; the rest keep the
 * manifest's order, which makes the list stable across builds.
 */
export function offeredFontFamilies(manifest: FontManifest): string[] {
    const offered = new Map<string, string>();
    for (const [family, entry] of Object.entries(manifest.families)) {
        if (undefined === entry) {
            continue;
        }

        offered.set(family, entry.substitutes[0] ?? family);
    }

    const fromDefault = offered.get(manifest.defaults.family);
    const rest = [...offered.entries()]
        .filter(([family]) => family !== manifest.defaults.family)
        .map(([, name]) => name);

    // De-duplicated: two families are free to answer to one name, and a select
    // showing it twice would be a bug the manifest cannot see.
    return [...new Set(undefined === fromDefault ? rest : [fromDefault, ...rest])];
}
