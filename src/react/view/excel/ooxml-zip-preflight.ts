export interface OoxmlZipPreflightLimits {
    maxPackageBytes: number;
    maxEntries: number;
    maxCentralDirectoryBytes: number;
    maxEntryUncompressedBytes: number;
    maxTotalUncompressedBytes: number;
    maxXmlEntryBytes: number;
    maxTotalXmlBytes: number;
}

export interface OoxmlZipEntryMetadata {
    name: string;
    compressedSize: number;
    uncompressedSize: number;
    directory: boolean;
    flags: number;
    compressionMethod: 0 | 8;
    localHeaderOffset: number;
    dataOffset: number;
}

export interface OoxmlZipPreflightResult {
    entries: Map<string, OoxmlZipEntryMetadata>;
    totalUncompressedBytes: number;
    totalXmlBytes: number;
}

export const DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS: Readonly<OoxmlZipPreflightLimits> = Object.freeze({
    maxPackageBytes: 512 * 1024 * 1024,
    maxEntries: 20_000,
    maxCentralDirectoryBytes: 64 * 1024 * 1024,
    maxEntryUncompressedBytes: 256 * 1024 * 1024,
    maxTotalUncompressedBytes: 512 * 1024 * 1024,
    maxXmlEntryBytes: 64 * 1024 * 1024,
    maxTotalXmlBytes: 256 * 1024 * 1024,
});

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;
const MAX_EOCD_SEARCH_BYTES = 65_557;

function fail(message: string): never {
    throw new Error(`Classeur OOXML refusé avant décompression : ${message}`);
}

function decodeEntryName(bytes: Uint8Array, utf8: boolean): string {
    if (!utf8 && bytes.some(byte => byte > 0x7f)) {
        return fail('un nom ZIP non ASCII sans indicateur UTF-8 est ambigu.');
    }
    try {
        return new TextDecoder(utf8 ? 'utf-8' : 'ascii', { fatal: true }).decode(bytes);
    } catch {
        return fail('un nom d’entrée ZIP n’est pas un texte valide.');
    }
}

function validateEntryName(name: string): { name: string; directory: boolean } {
    const directory = name.endsWith('/');
    const candidate = directory ? name.slice(0, -1) : name;
    if (
        !candidate
        || name.startsWith('/')
        || name.includes('\\')
        || name.includes('\0')
        || name.includes('//')
    ) {
        return fail(`nom d’entrée ZIP invalide : ${name.slice(0, 160)}`);
    }
    for (const segment of candidate.split('/')) {
        if (!segment || segment === '.' || segment === '..') {
            return fail(`nom d’entrée ZIP ambigu : ${name.slice(0, 160)}`);
        }
    }
    return { name, directory };
}

function isXmlMetadata(name: string): boolean {
    const normalized = name.toLowerCase();
    return normalized.endsWith('.xml') || normalized.endsWith('.rels');
}

function validateExtraFields(
    view: DataView,
    start: number,
    length: number,
    entryName: string,
): void {
    let cursor = start;
    const end = start + length;
    while (cursor < end) {
        if (cursor + 4 > end) return fail(`champ ZIP supplémentaire tronqué : ${entryName}.`);
        const headerId = view.getUint16(cursor, true);
        const dataSize = view.getUint16(cursor + 2, true);
        cursor += 4;
        if (cursor + dataSize > end) return fail(`champ ZIP supplémentaire mal formé : ${entryName}.`);
        // JSZip may replace the central filename with the Info-ZIP Unicode Path
        // field. Refuse that ambiguity so collision/path checks cover the exact
        // name the downstream parser will use. ZIP64 was already refused above.
        if (headerId === 0x0001 || headerId === 0x7075) {
            return fail(`champ ZIP de chemin ou taille ambigu non pris en charge : ${entryName}.`);
        }
        cursor += dataSize;
    }
}

export function inspectOoxmlZipCentralDirectory(
    source: ArrayBuffer | Uint8Array,
    limits: OoxmlZipPreflightLimits = DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS,
): OoxmlZipPreflightResult {
    const bytes = source instanceof Uint8Array
        ? source
        : new Uint8Array(source);
    if (bytes.byteLength < 22) return fail('archive ZIP incomplète.');
    if (bytes.byteLength > limits.maxPackageBytes) {
        return fail(`fichier supérieur à ${limits.maxPackageBytes} octets.`);
    }
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const firstCandidate = Math.max(0, bytes.byteLength - MAX_EOCD_SEARCH_BYTES);
    let eocdOffset = -1;
    for (let offset = bytes.byteLength - 22; offset >= firstCandidate; offset -= 1) {
        if (view.getUint32(offset, true) !== EOCD_SIGNATURE) continue;
        const commentLength = view.getUint16(offset + 20, true);
        if (offset + 22 + commentLength === bytes.byteLength) {
            eocdOffset = offset;
            break;
        }
    }
    if (eocdOffset < 0) return fail('fin de répertoire central introuvable.');

    const diskNumber = view.getUint16(eocdOffset + 4, true);
    const centralDisk = view.getUint16(eocdOffset + 6, true);
    const diskEntries = view.getUint16(eocdOffset + 8, true);
    const totalEntries = view.getUint16(eocdOffset + 10, true);
    const centralSize = view.getUint32(eocdOffset + 12, true);
    const centralOffset = view.getUint32(eocdOffset + 16, true);
    if (
        diskNumber !== 0
        || centralDisk !== 0
        || diskEntries !== totalEntries
        || totalEntries === 0xffff
        || centralSize === 0xffffffff
        || centralOffset === 0xffffffff
    ) {
        return fail('archives multi-disques et ZIP64 non prises en charge.');
    }
    if (totalEntries > limits.maxEntries) return fail(`plus de ${limits.maxEntries} entrées ZIP.`);
    if (centralSize > limits.maxCentralDirectoryBytes) return fail('répertoire central trop volumineux.');
    if (centralOffset + centralSize !== eocdOffset) return fail('bornes du répertoire central incohérentes.');

    const entries = new Map<string, OoxmlZipEntryMetadata>();
    const occupiedLocalRanges: Array<{ start: number; end: number; name: string }> = [];
    let cursor = centralOffset;
    let totalUncompressedBytes = 0;
    let totalXmlBytes = 0;
    for (let index = 0; index < totalEntries; index += 1) {
        if (cursor + 46 > eocdOffset || view.getUint32(cursor, true) !== CENTRAL_SIGNATURE) {
            return fail('répertoire central mal formé.');
        }
        const flags = view.getUint16(cursor + 8, true);
        const compressionMethod = view.getUint16(cursor + 10, true);
        const compressedSize = view.getUint32(cursor + 20, true);
        const uncompressedSize = view.getUint32(cursor + 24, true);
        const nameLength = view.getUint16(cursor + 28, true);
        const extraLength = view.getUint16(cursor + 30, true);
        const commentLength = view.getUint16(cursor + 32, true);
        const diskStart = view.getUint16(cursor + 34, true);
        const localHeaderOffset = view.getUint32(cursor + 42, true);
        const recordLength = 46 + nameLength + extraLength + commentLength;
        if (cursor + recordLength > eocdOffset) return fail('entrée du répertoire central tronquée.');
        if ((flags & 0x0001) !== 0) return fail('entrée ZIP chiffrée non prise en charge.');
        if (compressionMethod !== 0 && compressionMethod !== 8) {
            return fail(`méthode de compression ZIP non prise en charge : ${compressionMethod}.`);
        }
        if (
            compressedSize === 0xffffffff
            || uncompressedSize === 0xffffffff
            || localHeaderOffset === 0xffffffff
            || diskStart !== 0
        ) {
            return fail('métadonnées ZIP64 ou multi-disques non prises en charge.');
        }
        if (localHeaderOffset >= centralOffset) return fail('offset d’entrée locale incohérent.');
        const nameBytes = bytes.subarray(cursor + 46, cursor + 46 + nameLength);
        const decoded = decodeEntryName(nameBytes, (flags & 0x0800) !== 0);
        const validated = validateEntryName(decoded);
        validateExtraFields(view, cursor + 46 + nameLength, extraLength, decoded);
        const key = validated.name.toLowerCase();
        if (entries.has(key)) return fail(`entrée ZIP dupliquée ou collision de casse : ${decoded}.`);

        if (localHeaderOffset + 30 > centralOffset
            || view.getUint32(localHeaderOffset, true) !== LOCAL_SIGNATURE) {
            return fail(`en-tête local invalide : ${decoded}.`);
        }
        const localFlags = view.getUint16(localHeaderOffset + 6, true);
        const localCompressionMethod = view.getUint16(localHeaderOffset + 8, true);
        const localCompressedSize = view.getUint32(localHeaderOffset + 18, true);
        const localUncompressedSize = view.getUint32(localHeaderOffset + 22, true);
        const localNameLength = view.getUint16(localHeaderOffset + 26, true);
        const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
        const dataOffset = localHeaderOffset + 30 + localNameLength + localExtraLength;
        if (localFlags !== flags || localCompressionMethod !== compressionMethod) {
            return fail(`métadonnées locale et centrale divergentes : ${decoded}.`);
        }
        if (dataOffset > centralOffset || dataOffset + compressedSize > centralOffset) {
            return fail(`données compressées hors limites : ${decoded}.`);
        }
        const localNameBytes = bytes.subarray(
            localHeaderOffset + 30,
            localHeaderOffset + 30 + localNameLength,
        );
        const localName = decodeEntryName(localNameBytes, (localFlags & 0x0800) !== 0);
        if (localName !== decoded) return fail(`nom local et central divergent : ${decoded}.`);
        validateExtraFields(
            view,
            localHeaderOffset + 30 + localNameLength,
            localExtraLength,
            decoded,
        );
        if ((flags & 0x0008) === 0 && (
            localCompressedSize !== compressedSize
            || localUncompressedSize !== uncompressedSize
        )) {
            return fail(`tailles locale et centrale divergentes : ${decoded}.`);
        }
        occupiedLocalRanges.push({
            start: localHeaderOffset,
            end: dataOffset + compressedSize,
            name: decoded,
        });

        if (!validated.directory) {
            if (uncompressedSize > limits.maxEntryUncompressedBytes) {
                return fail(`entrée trop volumineuse : ${decoded}.`);
            }
            totalUncompressedBytes += uncompressedSize;
            if (totalUncompressedBytes > limits.maxTotalUncompressedBytes) {
                return fail('taille totale décompressée supérieure à la limite.');
            }
            if (isXmlMetadata(decoded)) {
                if (uncompressedSize > limits.maxXmlEntryBytes) {
                    return fail(`partie XML trop volumineuse : ${decoded}.`);
                }
                totalXmlBytes += uncompressedSize;
                if (totalXmlBytes > limits.maxTotalXmlBytes) {
                    return fail('taille XML décompressée totale supérieure à la limite.');
                }
            }
        }
        entries.set(key, {
            name: validated.name,
            compressedSize,
            uncompressedSize,
            directory: validated.directory,
            flags,
            compressionMethod: compressionMethod as 0 | 8,
            localHeaderOffset,
            dataOffset,
        });
        cursor += recordLength;
    }
    if (cursor !== eocdOffset) return fail('données non analysées dans le répertoire central.');
    occupiedLocalRanges.sort((left, right) => left.start - right.start);
    for (let index = 1; index < occupiedLocalRanges.length; index += 1) {
        const previous = occupiedLocalRanges[index - 1];
        const current = occupiedLocalRanges[index];
        if (current.start < previous.end) {
            return fail(`entrées ZIP locales superposées : ${previous.name} et ${current.name}.`);
        }
    }
    return { entries, totalUncompressedBytes, totalXmlBytes };
}

async function countDeflatedBytes(
    compressed: Uint8Array,
    entry: OoxmlZipEntryMetadata,
    limits: OoxmlZipPreflightLimits,
    remainingTotalBytes: number,
    remainingXmlBytes: number,
): Promise<number> {
    if (typeof DecompressionStream !== 'function') {
        return fail('le moteur ne permet pas de vérifier la décompression DEFLATE de façon bornée.');
    }
    let stream: DecompressionStream;
    try {
        stream = new DecompressionStream('deflate-raw');
    } catch {
        return fail('le moteur ne prend pas en charge DEFLATE brut de façon bornée.');
    }
    const reader = new Blob([compressed]).stream().pipeThrough(stream).getReader();
    let produced = 0;
    try {
        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            produced += value.byteLength;
            if (produced > entry.uncompressedSize) {
                return fail(`taille décompressée réelle supérieure à la taille déclarée : ${entry.name}.`);
            }
            if (produced > limits.maxEntryUncompressedBytes || produced > remainingTotalBytes) {
                return fail(`budget de décompression dépassé : ${entry.name}.`);
            }
            if (isXmlMetadata(entry.name)
                && (produced > limits.maxXmlEntryBytes || produced > remainingXmlBytes)) {
                return fail(`budget XML réel dépassé : ${entry.name}.`);
            }
        }
    } catch (error) {
        if (error instanceof Error && error.message.startsWith('Classeur OOXML refusé')) throw error;
        return fail(`flux DEFLATE invalide : ${entry.name}.`);
    } finally {
        void reader.cancel().catch(() => undefined);
    }
    return produced;
}

/**
 * Streams every compressed member once before JSZip/ExcelJS can inflate it.
 * This independently verifies the real produced byte count instead of trusting
 * attacker-controlled central-directory sizes.
 */
export async function validateOoxmlZipInflationBounds(
    source: ArrayBuffer | Uint8Array,
    inspected: OoxmlZipPreflightResult,
    limits: OoxmlZipPreflightLimits = DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS,
): Promise<void> {
    const bytes = source instanceof Uint8Array ? source : new Uint8Array(source);
    let totalBytes = 0;
    let totalXmlBytes = 0;
    for (const entry of inspected.entries.values()) {
        if (entry.directory) {
            if (entry.compressedSize !== 0 || entry.uncompressedSize !== 0) {
                return fail(`répertoire ZIP avec contenu : ${entry.name}.`);
            }
            continue;
        }
        const compressed = bytes.subarray(
            entry.dataOffset,
            entry.dataOffset + entry.compressedSize,
        );
        let produced: number;
        if (entry.compressionMethod === 0) {
            produced = compressed.byteLength;
        } else {
            produced = await countDeflatedBytes(
                compressed,
                entry,
                limits,
                limits.maxTotalUncompressedBytes - totalBytes,
                limits.maxTotalXmlBytes - totalXmlBytes,
            );
        }
        if (produced !== entry.uncompressedSize) {
            return fail(`taille décompressée réelle différente de la taille déclarée : ${entry.name}.`);
        }
        totalBytes += produced;
        if (totalBytes > limits.maxTotalUncompressedBytes) {
            return fail('taille totale réellement décompressée supérieure à la limite.');
        }
        if (isXmlMetadata(entry.name)) {
            totalXmlBytes += produced;
            if (totalXmlBytes > limits.maxTotalXmlBytes) {
                return fail('taille XML réellement décompressée supérieure à la limite.');
            }
        }
    }
}
