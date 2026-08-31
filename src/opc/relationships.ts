import { XmlDocument } from '../ooxml/xml.js';

/**
 * OPC relationships — how a package says which part is which.
 *
 * `word/document.xml` is a convention, not a rule. The main part is whatever
 * the package relationship of type `.../officeDocument` points at, and styles
 * are whatever that part's own relationships point at. Producers do deviate:
 * the name is theirs to choose, and a reader that hard-codes the conventional
 * path opens most files and mysteriously fails on the rest.
 *
 * Following the relationships is barely more work and is what the format
 * actually specifies.
 */

export interface Relationship {
    readonly id: string;
    readonly type: string;
    readonly target: string;
    /** "External" targets are URLs, not parts — a hyperlink, not a file. */
    readonly external: boolean;
}

const OFFICE_DOCUMENT = 'officeDocument';

export function readRelationships(xml: string): Relationship[] {
    const root = XmlDocument.parse(xml).root;

    return root.elements('Relationship').map((element) => ({
        id: element.attribute('Id') ?? '',
        type: element.attribute('Type') ?? '',
        target: element.attribute('Target') ?? '',
        external: 'External' === element.attribute('TargetMode'),
    }));
}

/** The part holding a part's relationships: `a/b.xml` to `a/_rels/b.xml.rels`. */
export function relationshipPartFor(part: string): string {
    const cut = part.lastIndexOf('/');
    const folder = cut < 0 ? '' : part.slice(0, cut + 1);
    const file = cut < 0 ? part : part.slice(cut + 1);

    return `${folder}_rels/${file}.rels`;
}

/**
 * Resolve a relationship target against the part that declared it.
 *
 * Targets are relative to the source part's FOLDER, not to the package root —
 * `styles.xml` declared by `word/document.xml` means `word/styles.xml`. A
 * leading slash makes it absolute from the root instead.
 */
export function resolveTarget(sourcePart: string, target: string): string {
    if (target.startsWith('/')) {
        return target.slice(1);
    }

    const cut = sourcePart.lastIndexOf('/');
    const folder = cut < 0 ? '' : sourcePart.slice(0, cut + 1);
    const segments: string[] = [];

    for (const segment of `${folder}${target}`.split('/')) {
        if ('.' === segment || '' === segment) {
            continue;
        }
        if ('..' === segment) {
            segments.pop();
            continue;
        }
        segments.push(segment);
    }

    return segments.join('/');
}

/**
 * Find a relationship by the LAST segment of its type.
 *
 * Matching on the tail rather than the whole URI because the namespace differs
 * between the strict and transitional flavours of the format — the same
 * relationship, two URIs, and a full-string match silently finds neither on a
 * strict document.
 */
export function findByType(relationships: readonly Relationship[], type: string): Relationship | null {
    for (const relationship of relationships) {
        if (!relationship.external && relationship.type.split('/').pop() === type) {
            return relationship;
        }
    }

    return null;
}

/** A relationship by its Id, which is how a `w:headerReference` names one. */
export function findById(relationships: readonly Relationship[], id: string): Relationship | null {
    for (const relationship of relationships) {
        if (!relationship.external && relationship.id === id) {
            return relationship;
        }
    }

    return null;
}

export { OFFICE_DOCUMENT };
