const SPREADSHEETML_MAIN_NAMESPACE =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

type MarkupVisitor = (tagStart: number, tagEnd: number, closing: boolean) => void;

const findTagEnd = (xml: string, start: number): number => {
    let quote = '';
    for (let index = start; index < xml.length; index += 1) {
        const character = xml[index];
        if (quote) {
            if (character === quote) quote = '';
        } else if (character === '"' || character === '\'') {
            quote = character;
        } else if (character === '>') {
            return index;
        }
    }
    return -1;
};

const scanMarkup = (xml: string, visitor: MarkupVisitor): boolean => {
    let cursor = 0;
    while (cursor < xml.length) {
        const tagStart = xml.indexOf('<', cursor);
        if (tagStart < 0) return true;

        if (xml.startsWith('<!--', tagStart)) {
            const end = xml.indexOf('-->', tagStart + 4);
            if (end < 0) return false;
            cursor = end + 3;
            continue;
        }
        if (xml.startsWith('<![CDATA[', tagStart)) {
            const end = xml.indexOf(']]>', tagStart + 9);
            if (end < 0) return false;
            cursor = end + 3;
            continue;
        }
        if (xml.startsWith('<?', tagStart)) {
            const end = xml.indexOf('?>', tagStart + 2);
            if (end < 0) return false;
            cursor = end + 2;
            continue;
        }

        const tagEnd = findTagEnd(xml, tagStart + 1);
        if (tagEnd < 0) return false;
        if (!xml.startsWith('<!', tagStart)) {
            visitor(tagStart, tagEnd, xml[tagStart + 1] === '/');
        }
        cursor = tagEnd + 1;
    }
    return true;
};

export function normalizeSpreadsheetMlElementPrefixes(
    xml: string,
): { xml: string; changed: boolean } {
    const namespaceBindings = new Map<string, Set<string>>();
    const wellFormed = scanMarkup(xml, (tagStart, tagEnd, closing) => {
        if (closing) return;
        const tag = xml.slice(tagStart, tagEnd + 1);
        const declaration =
            /\sxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])([^"']*)\2/g;
        for (const match of tag.matchAll(declaration)) {
            const bindings = namespaceBindings.get(match[1]) ?? new Set<string>();
            bindings.add(match[3]);
            namespaceBindings.set(match[1], bindings);
        }
    });
    if (!wellFormed) return { xml, changed: false };

    const prefixes = new Set(
        [...namespaceBindings]
            .filter(([, bindings]) =>
                bindings.size === 1 && bindings.has(SPREADSHEETML_MAIN_NAMESPACE))
            .map(([prefix]) => prefix),
    );
    if (prefixes.size === 0) return { xml, changed: false };

    const pieces: string[] = [];
    let copyStart = 0;
    let changed = false;
    const secondPassWellFormed = scanMarkup(xml, (tagStart, _tagEnd, closing) => {
        const nameStart = tagStart + (closing ? 2 : 1);
        if (!/[A-Za-z_]/.test(xml[nameStart] ?? '')) return;

        let nameEnd = nameStart + 1;
        while (/[A-Za-z0-9_.:-]/.test(xml[nameEnd] ?? '')) nameEnd += 1;
        const qualifiedName = xml.slice(nameStart, nameEnd);
        const colon = qualifiedName.indexOf(':');
        if (colon <= 0 || !prefixes.has(qualifiedName.slice(0, colon))) return;

        pieces.push(xml.slice(copyStart, nameStart));
        pieces.push(qualifiedName.slice(colon + 1));
        copyStart = nameEnd;
        changed = true;
    });
    if (!secondPassWellFormed || !changed) return { xml, changed: false };

    pieces.push(xml.slice(copyStart));
    return { xml: pieces.join(''), changed: true };
}

const relationshipBaseSegments = (relationshipPath: string): string[] | null => {
    if (relationshipPath.toLowerCase() === '_rels/.rels') return [];
    const match = /^(.*)\/_rels\/[^/]+\.rels$/i.exec(relationshipPath);
    return match ? match[1].split('/') : null;
};

const makeRelativePackageTarget = (
    target: string,
    relationshipPath: string,
): string | null => {
    const baseSegments = relationshipBaseSegments(relationshipPath);
    if (!baseSegments || !target.startsWith('/xl/') || target.startsWith('//')) {
        return null;
    }

    const suffixStart = target.search(/[?#]/);
    const path = suffixStart < 0 ? target : target.slice(0, suffixStart);
    const suffix = suffixStart < 0 ? '' : target.slice(suffixStart);
    const targetSegments = path.slice(1).split('/');
    if (targetSegments.some((segment) =>
        !segment || segment === '.' || segment === '..')) {
        return null;
    }

    let commonLength = 0;
    while (
        commonLength < baseSegments.length
        && commonLength < targetSegments.length
        && baseSegments[commonLength] === targetSegments[commonLength]
    ) {
        commonLength += 1;
    }
    const relativeSegments = [
        ...Array(baseSegments.length - commonLength).fill('..'),
        ...targetSegments.slice(commonLength),
    ];
    return relativeSegments.length > 0
        ? `${relativeSegments.join('/')}${suffix}`
        : null;
};

export function normalizeOoxmlRelationshipTargets(
    xml: string,
    relationshipPath: string,
): { xml: string; changed: boolean } {
    if (
        !xml.includes(
            'http://schemas.openxmlformats.org/package/2006/relationships',
        )
        || relationshipBaseSegments(relationshipPath) === null
    ) {
        return { xml, changed: false };
    }

    const pieces: string[] = [];
    let copyStart = 0;
    let changed = false;
    const wellFormed = scanMarkup(xml, (tagStart, tagEnd, closing) => {
        if (closing) return;
        const tag = xml.slice(tagStart, tagEnd + 1);
        const name = /^<([A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)/.exec(tag);
        if (name?.[2] !== 'Relationship') return;
        if (/\sTargetMode\s*=\s*(["'])External\1/i.test(tag)) return;

        const target = /\sTarget\s*=\s*(["'])(.*?)\1/.exec(tag);
        if (!target) return;
        const relative = makeRelativePackageTarget(target[2], relationshipPath);
        if (!relative) return;

        const valueStart = tagStart + target.index + target[0].indexOf(target[2]);
        pieces.push(xml.slice(copyStart, valueStart));
        pieces.push(relative);
        copyStart = valueStart + target[2].length;
        changed = true;
    });
    if (!wellFormed || !changed) return { xml, changed: false };

    pieces.push(xml.slice(copyStart));
    return { xml: pieces.join(''), changed: true };
}
