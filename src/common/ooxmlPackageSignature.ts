import { open, type FileHandle } from 'fs/promises';
import { extname, posix } from 'path';
import { inflateRaw } from 'zlib';

const MAX_ZIP_ENTRIES = 20_000;
const MAX_CENTRAL_DIRECTORY_BYTES = 64 * 1024 * 1024;
const MAX_METADATA_PART_BYTES = 1024 * 1024;
const MAX_COMPRESSED_METADATA_PART_BYTES = 2 * 1024 * 1024;
const MAX_RELATIONSHIP_PARTS = 4_096;
const MAX_RELATIONSHIP_XML_BYTES = 16 * 1024 * 1024;
export const MAX_VIRTUAL_OOXML_PACKAGE_BYTES = 128 * 1024 * 1024;

const CONTENT_TYPES_PART = '/[Content_Types].xml';
const ROOT_RELATIONSHIPS_PART = '/_rels/.rels';
const RELATIONSHIPS_NAMESPACE =
    'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES_NAMESPACE =
    'http://schemas.openxmlformats.org/package/2006/content-types';
const ORIGIN_RELATIONSHIP_TYPE =
    'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin';
const SIGNATURE_RELATIONSHIP_TYPE =
    'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature';
const ORIGIN_CONTENT_TYPE =
    'application/vnd.openxmlformats-package.digital-signature-origin';
const SIGNATURE_CONTENT_TYPE =
    'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml';
const XLM_MACRO_SHEET_CONTENT_TYPES = new Set([
    'application/vnd.ms-excel.macrosheet+xml',
    'application/vnd.ms-excel.intlmacrosheet+xml',
    'application/vnd.ms-excel.macrosheet',
    'application/vnd.ms-excel.intlmacrosheet',
]);
const XLM_MACRO_SHEET_RELATIONSHIP_TYPES = new Set([
    'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet',
    'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet',
]);

export const OOXML_XLM_AUTOMATION_BLOCKED_MESSAGE =
    'Ouverture automatisée refusée : ce classeur contient une feuille macro Excel 4.0 (XLM), que le mode AutomationSecurity ne désactive pas.';
export const OOXML_XLM_VERIFICATION_BLOCKED_MESSAGE =
    'Ouverture automatisée refusée : l’absence de feuilles macro Excel 4.0 (XLM) n’a pas pu être vérifiée de façon sûre.';

export const OOXML_PACKAGE_SIGNATURE_WRITE_BLOCKED_MESSAGE =
    'Écriture refusée : ce classeur porte une signature numérique de package Office. Toute modification invaliderait cette signature.';
export const OOXML_PACKAGE_SIGNATURE_VERIFICATION_BLOCKED_MESSAGE =
    'Écriture refusée : l’état de la signature numérique de package Office n’a pas pu être vérifié de façon sûre.';

const OOXML_SPREADSHEET_EXTENSIONS = new Set([
    '.xlsx',
    '.xlsm',
    '.xlsb',
    '.xltx',
    '.xltm',
    '.xlam',
]);

interface RandomAccessReader {
    readonly size: number;
    read(offset: number, length: number): Promise<Buffer>;
}

interface ZipEntry {
    readonly partName: string;
    readonly compressionMethod: number;
    readonly flags: number;
    readonly crc32: number;
    readonly compressedSize: number;
    readonly uncompressedSize: number;
    readonly localHeaderOffset: number;
}

interface ZipPackage {
    readonly reader: RandomAccessReader;
    readonly entries: Map<string, ZipEntry>;
}

interface XmlElement {
    readonly localName: string;
    readonly namespaceUri: string;
    readonly attributes: Map<string, string>;
    readonly depth: number;
}

interface ContentTypeTable {
    readonly defaults: Map<string, string>;
    readonly overrides: Map<string, string>;
}

interface OpcRelationship {
    readonly id: string;
    readonly type: string;
    readonly target: string;
    readonly targetMode?: string;
}

const CRC32_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let index = 0; index < 256; index += 1) {
        let value = index;
        for (let bit = 0; bit < 8; bit += 1) {
            value = (value & 1) !== 0
                ? 0xedb88320 ^ (value >>> 1)
                : value >>> 1;
        }
        table[index] = value >>> 0;
    }
    return table;
})();

function calculateCrc32(value: Buffer): number {
    let crc = 0xffffffff;
    for (const octet of value) {
        crc = CRC32_TABLE[(crc ^ octet) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
}

function fail(reason: string): never {
    throw new Error(`Office package signature verification failed: ${reason}`);
}

function inflateRawBounded(value: Buffer, maximumBytes: number): Promise<Buffer> {
    return new Promise((resolve, reject) => {
        inflateRaw(value, { maxOutputLength: maximumBytes }, (error, result) => {
            if (error) {
                reject(error);
                return;
            }
            resolve(result);
        });
    });
}

function lowerAscii(value: string): string {
    return value.replace(/[A-Z]/g, (character) => character.toLowerCase());
}

function decodeZipName(value: Buffer, utf8: boolean): string {
    if (!utf8 && value.some((octet) => octet > 0x7f)) {
        return fail('a legacy non-ASCII ZIP entry name is ambiguous');
    }
    try {
        return new TextDecoder(utf8 ? 'utf-8' : 'ascii', { fatal: true }).decode(value);
    } catch {
        return fail('a ZIP entry name is not valid text');
    }
}

function validatePercentEncoding(value: string): void {
    for (let index = 0; index < value.length; index += 1) {
        if (value[index] !== '%') {
            continue;
        }
        const encoded = value.slice(index + 1, index + 3);
        if (!/^[0-9A-Fa-f]{2}$/.test(encoded)) {
            fail('a package URI contains invalid percent encoding');
        }
        const octet = Number.parseInt(encoded, 16);
        if (octet === 0x2f || octet === 0x5c || octet === 0x00) {
            fail('a package URI contains an encoded path separator or NUL');
        }
        const character = String.fromCharCode(octet);
        if (/^[A-Za-z0-9._~-]$/.test(character)) {
            fail('a package URI percent-encodes an unreserved character');
        }
        index += 2;
    }
}

function validatePartName(partName: string): string {
    if (
        !partName.startsWith('/') ||
        partName.length < 2 ||
        partName.endsWith('/') ||
        partName.includes('\\') ||
        partName.includes('\0') ||
        partName.includes('?') ||
        partName.includes('#') ||
        partName.includes('//')
    ) {
        return fail(`invalid OPC part name: ${partName}`);
    }
    validatePercentEncoding(partName);
    for (const segment of partName.slice(1).split('/')) {
        if (segment === '' || segment === '.' || segment === '..') {
            return fail(`invalid OPC part name: ${partName}`);
        }
    }
    return partName;
}

function validateZipEntryName(name: string): string | undefined {
    if (name.endsWith('/')) {
        const directoryPart = name.slice(0, -1);
        if (directoryPart.length === 0) {
            fail('the ZIP contains an invalid root directory entry');
        }
        validatePartName(`/${directoryPart}`);
        return undefined;
    }
    if (name.startsWith('/')) {
        fail('a ZIP entry name starts with a slash');
    }
    return validatePartName(`/${name}`);
}

function createBufferReader(value: Uint8Array): RandomAccessReader {
    if (value.byteLength > MAX_VIRTUAL_OOXML_PACKAGE_BYTES) {
        fail(
            `virtual package exceeds the ${MAX_VIRTUAL_OOXML_PACKAGE_BYTES}-byte inspection limit`,
        );
    }
    const data = Buffer.from(value.buffer, value.byteOffset, value.byteLength);
    return {
        size: data.length,
        async read(offset, length) {
            if (offset < 0 || length < 0 || offset + length > data.length) {
                return fail('ZIP read is outside the package bounds');
            }
            return data.subarray(offset, offset + length);
        },
    };
}

function createFileReader(handle: FileHandle, size: number): RandomAccessReader {
    return {
        size,
        async read(offset, length) {
            if (offset < 0 || length < 0 || offset + length > size) {
                return fail('ZIP read is outside the package bounds');
            }
            const result = Buffer.allocUnsafe(length);
            const { bytesRead } = await handle.read(result, 0, length, offset);
            if (bytesRead !== length) {
                return fail('ZIP package ended during inspection');
            }
            return result;
        },
    };
}

async function openZipPackage(reader: RandomAccessReader): Promise<ZipPackage> {
    if (reader.size < 22) {
        return fail('the file is not a complete ZIP package');
    }
    const tailLength = Math.min(reader.size, 65_557);
    const tailOffset = reader.size - tailLength;
    const tail = await reader.read(tailOffset, tailLength);
    let eocdOffset = -1;
    for (let offset = tail.length - 22; offset >= 0; offset -= 1) {
        if (
            tail.readUInt32LE(offset) === 0x06054b50 &&
            offset + 22 + tail.readUInt16LE(offset + 20) === tail.length
        ) {
            eocdOffset = offset;
            break;
        }
    }
    if (eocdOffset < 0) {
        return fail('ZIP end-of-central-directory record was not found');
    }
    const diskNumber = tail.readUInt16LE(eocdOffset + 4);
    const centralDisk = tail.readUInt16LE(eocdOffset + 6);
    const diskEntries = tail.readUInt16LE(eocdOffset + 8);
    const totalEntries = tail.readUInt16LE(eocdOffset + 10);
    const centralSize = tail.readUInt32LE(eocdOffset + 12);
    const centralOffset = tail.readUInt32LE(eocdOffset + 16);
    if (
        diskNumber !== 0 ||
        centralDisk !== 0 ||
        diskEntries !== totalEntries ||
        totalEntries === 0xffff ||
        centralSize === 0xffffffff ||
        centralOffset === 0xffffffff
    ) {
        return fail('multi-disk or ZIP64 packages are not supported by the bounded verifier');
    }
    if (totalEntries > MAX_ZIP_ENTRIES) {
        return fail(`ZIP contains more than ${MAX_ZIP_ENTRIES} entries`);
    }
    if (centralSize > MAX_CENTRAL_DIRECTORY_BYTES) {
        return fail('ZIP central directory exceeds the inspection limit');
    }
    const absoluteEocdOffset = tailOffset + eocdOffset;
    if (centralOffset + centralSize !== absoluteEocdOffset) {
        return fail('ZIP central-directory bounds are inconsistent');
    }
    const central = await reader.read(centralOffset, centralSize);
    const entries = new Map<string, ZipEntry>();
    let cursor = 0;
    for (let index = 0; index < totalEntries; index += 1) {
        if (cursor + 46 > central.length || central.readUInt32LE(cursor) !== 0x02014b50) {
            return fail('ZIP central directory is malformed');
        }
        const flags = central.readUInt16LE(cursor + 8);
        const compressionMethod = central.readUInt16LE(cursor + 10);
        const crc32 = central.readUInt32LE(cursor + 16);
        const compressedSize = central.readUInt32LE(cursor + 20);
        const uncompressedSize = central.readUInt32LE(cursor + 24);
        const nameLength = central.readUInt16LE(cursor + 28);
        const extraLength = central.readUInt16LE(cursor + 30);
        const commentLength = central.readUInt16LE(cursor + 32);
        const diskStart = central.readUInt16LE(cursor + 34);
        const localHeaderOffset = central.readUInt32LE(cursor + 42);
        const recordLength = 46 + nameLength + extraLength + commentLength;
        if (cursor + recordLength > central.length) {
            return fail('ZIP central-directory entry is truncated');
        }
        if (
            diskStart !== 0 ||
            compressedSize === 0xffffffff ||
            uncompressedSize === 0xffffffff ||
            localHeaderOffset === 0xffffffff
        ) {
            return fail('ZIP64 entry metadata is not supported by the bounded verifier');
        }
        if ((flags & 0x0001) !== 0) {
            return fail('encrypted ZIP entries cannot be inspected safely');
        }
        const name = decodeZipName(
            central.subarray(cursor + 46, cursor + 46 + nameLength),
            (flags & 0x0800) !== 0,
        );
        const partName = validateZipEntryName(name);
        if (partName) {
            const key = lowerAscii(partName);
            if (entries.has(key)) {
                return fail(`duplicate or case-colliding ZIP part: ${partName}`);
            }
            entries.set(key, {
                partName,
                compressionMethod,
                flags,
                crc32,
                compressedSize,
                uncompressedSize,
                localHeaderOffset,
            });
        }
        cursor += recordLength;
    }
    if (cursor !== central.length) {
        return fail('ZIP central directory contains unparsed data');
    }
    return { reader, entries };
}

function getEntry(zip: ZipPackage, partName: string): ZipEntry | undefined {
    return zip.entries.get(lowerAscii(validatePartName(partName)));
}

async function readEntry(
    zip: ZipPackage,
    entry: ZipEntry,
    maximumBytes = MAX_METADATA_PART_BYTES,
): Promise<Buffer> {
    if (entry.uncompressedSize > maximumBytes) {
        return fail(`${entry.partName} exceeds the metadata inspection limit`);
    }
    if (entry.compressedSize > MAX_COMPRESSED_METADATA_PART_BYTES) {
        return fail(`${entry.partName} compressed data exceeds the inspection limit`);
    }
    const header = await zip.reader.read(entry.localHeaderOffset, 30);
    if (header.readUInt32LE(0) !== 0x04034b50) {
        return fail(`local ZIP header is missing for ${entry.partName}`);
    }
    const localFlags = header.readUInt16LE(6);
    const localMethod = header.readUInt16LE(8);
    const localCrc32 = header.readUInt32LE(14);
    const localCompressedSize = header.readUInt32LE(18);
    const localUncompressedSize = header.readUInt32LE(22);
    const nameLength = header.readUInt16LE(26);
    const extraLength = header.readUInt16LE(28);
    if (localFlags !== entry.flags || localMethod !== entry.compressionMethod) {
        return fail(`local and central ZIP metadata disagree for ${entry.partName}`);
    }
    if (
        (localFlags & 0x0008) === 0 &&
        (
            localCrc32 !== entry.crc32 ||
            localCompressedSize !== entry.compressedSize ||
            localUncompressedSize !== entry.uncompressedSize
        )
    ) {
        return fail(`local and central ZIP sizes or CRC disagree for ${entry.partName}`);
    }
    const localNameBytes = await zip.reader.read(entry.localHeaderOffset + 30, nameLength);
    const localName = decodeZipName(localNameBytes, (localFlags & 0x0800) !== 0);
    if (`/${localName}` !== entry.partName) {
        return fail(`local and central ZIP names disagree for ${entry.partName}`);
    }
    const dataOffset = entry.localHeaderOffset + 30 + nameLength + extraLength;
    const compressed = await zip.reader.read(dataOffset, entry.compressedSize);
    let result: Buffer;
    if (entry.compressionMethod === 0) {
        result = compressed;
    } else if (entry.compressionMethod === 8) {
        try {
            result = await inflateRawBounded(compressed, maximumBytes);
        } catch {
            return fail(`compressed metadata could not be inflated for ${entry.partName}`);
        }
    } else {
        return fail(`unsupported ZIP compression method for ${entry.partName}`);
    }
    if (
        result.length !== entry.uncompressedSize ||
        calculateCrc32(result) !== entry.crc32
    ) {
        return fail(`ZIP size or CRC is invalid for ${entry.partName}`);
    }
    return result;
}

function decodeXmlEntity(value: string): string {
    return value.replace(/&(#x[0-9A-Fa-f]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (match, entity) => {
        if (entity === 'amp') return '&';
        if (entity === 'lt') return '<';
        if (entity === 'gt') return '>';
        if (entity === 'quot') return '"';
        if (entity === 'apos') return "'";
        const codePoint = entity.startsWith('#x')
            ? Number.parseInt(entity.slice(2), 16)
            : Number.parseInt(entity.slice(1), 10);
        if (!Number.isSafeInteger(codePoint) || codePoint <= 0 || codePoint > 0x10ffff) {
            return fail('XML contains an invalid character reference');
        }
        return String.fromCodePoint(codePoint);
    });
}

function decodeXmlAttribute(value: string): string {
    if (/&(?!(?:#x[0-9A-Fa-f]+|#[0-9]+|amp|lt|gt|quot|apos);)/.test(value)) {
        fail('XML contains an unknown or unterminated entity');
    }
    return decodeXmlEntity(value);
}

function splitQualifiedName(value: string): { prefix: string; localName: string } {
    const separator = value.indexOf(':');
    if (separator < 0) {
        return { prefix: '', localName: value };
    }
    if (separator === 0 || separator === value.length - 1 || value.indexOf(':', separator + 1) >= 0) {
        return fail('XML contains an invalid qualified name');
    }
    return { prefix: value.slice(0, separator), localName: value.slice(separator + 1) };
}

function decodeOpcXml(value: Buffer): string {
    if (value.length === 0) return fail('metadata XML is empty');
    let encoding: 'utf-8' | 'utf-16le' | 'utf-16be' = 'utf-8';
    let offset = 0;
    if (
        value.length >= 4 &&
        (
            (value[0] === 0x00 && value[1] === 0x00 && value[2] === 0xfe && value[3] === 0xff) ||
            (value[0] === 0xff && value[1] === 0xfe && value[2] === 0x00 && value[3] === 0x00)
        )
    ) {
        return fail('UTF-32 OPC metadata is not supported by the bounded verifier');
    }
    if (value.length >= 3 && value[0] === 0xef && value[1] === 0xbb && value[2] === 0xbf) {
        encoding = 'utf-8';
        offset = 3;
    } else if (value.length >= 2 && value[0] === 0xff && value[1] === 0xfe) {
        encoding = 'utf-16le';
        offset = 2;
    } else if (value.length >= 2 && value[0] === 0xfe && value[1] === 0xff) {
        encoding = 'utf-16be';
        offset = 2;
    } else if (value.length >= 4 && value[0] === 0x3c && value[1] === 0x00) {
        encoding = 'utf-16le';
    } else if (value.length >= 4 && value[0] === 0x00 && value[1] === 0x3c) {
        encoding = 'utf-16be';
    }
    const encoded = value.subarray(offset);
    if (encoding !== 'utf-8' && encoded.length % 2 !== 0) {
        return fail('UTF-16 OPC metadata has an odd byte length');
    }
    try {
        const xml = new TextDecoder(encoding, { fatal: true }).decode(encoded);
        if (
            xml.includes('\u0000') ||
            /[\u0001-\u0008\u000B\u000C\u000E-\u001F]/.test(xml)
        ) {
            return fail('metadata XML contains forbidden control characters');
        }
        return xml.charCodeAt(0) === 0xfeff ? xml.slice(1) : xml;
    } catch {
        return fail(`metadata XML is not valid ${encoding.toUpperCase()}`);
    }
}

function parseXmlElements(value: Buffer): XmlElement[] {
    let xml: string;
    xml = decodeOpcXml(value);
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) {
        return fail('DTD and entity declarations are forbidden in package metadata');
    }
    const elements: XmlElement[] = [];
    const stack: Array<{ qName: string; namespaces: Map<string, string> }> = [];
    let cursor = 0;
    let rootClosed = false;
    while (cursor < xml.length) {
        const open = xml.indexOf('<', cursor);
        if (open < 0) {
            if (xml.slice(cursor).trim() !== '') fail('XML has text outside its root element');
            break;
        }
        if (xml.slice(cursor, open).trim() !== '') {
            fail('package metadata XML contains unexpected text');
        }
        if (xml.startsWith('<!--', open)) {
            const close = xml.indexOf('-->', open + 4);
            if (close < 0) fail('XML comment is unterminated');
            cursor = close + 3;
            continue;
        }
        if (xml.startsWith('<?', open)) {
            const close = xml.indexOf('?>', open + 2);
            if (close < 0) fail('XML processing instruction is unterminated');
            cursor = close + 2;
            continue;
        }
        if (xml.startsWith('<!', open)) {
            return fail('unsupported XML declaration in package metadata');
        }
        const close = xml.indexOf('>', open + 1);
        if (close < 0) return fail('XML tag is unterminated');
        let token = xml.slice(open + 1, close);
        if (token.includes('<')) return fail('XML tag is malformed');
        if (token.startsWith('/')) {
            const qName = token.slice(1).trim();
            if (!/^[A-Za-z_][A-Za-z0-9_.:-]*$/.test(qName)) {
                return fail('XML closing tag is malformed');
            }
            const current = stack.pop();
            if (!current || current.qName !== qName) return fail('XML tags are not balanced');
            if (stack.length === 0) rootClosed = true;
            cursor = close + 1;
            continue;
        }
        const selfClosing = /\/\s*$/.test(token);
        if (selfClosing) token = token.replace(/\/\s*$/, '');
        const nameMatch = /^\s*([A-Za-z_][A-Za-z0-9_.:-]*)/.exec(token);
        if (!nameMatch) return fail('XML start tag is malformed');
        if (rootClosed && stack.length === 0) return fail('XML contains multiple root elements');
        const qName = nameMatch[1];
        let attributeText = token.slice(nameMatch[0].length);
        const rawAttributes = new Map<string, string>();
        while (attributeText.trim() !== '') {
            const attributeMatch = /^\s+([A-Za-z_][A-Za-z0-9_.:-]*)\s*=\s*(["'])([\s\S]*?)\2/.exec(attributeText);
            if (!attributeMatch) return fail('XML attribute syntax is malformed');
            const attributeName = attributeMatch[1];
            if (rawAttributes.has(attributeName)) return fail('XML contains duplicate attributes');
            rawAttributes.set(attributeName, decodeXmlAttribute(attributeMatch[3]));
            attributeText = attributeText.slice(attributeMatch[0].length);
        }
        const namespaces = new Map(stack.at(-1)?.namespaces ?? []);
        for (const [attributeName, attributeValue] of rawAttributes) {
            if (attributeName === 'xmlns') namespaces.set('', attributeValue);
            else if (attributeName.startsWith('xmlns:')) namespaces.set(attributeName.slice(6), attributeValue);
        }
        const { prefix, localName } = splitQualifiedName(qName);
        const namespaceUri = namespaces.get(prefix);
        if (namespaceUri === undefined) return fail(`XML prefix is not declared: ${prefix}`);
        const attributes = new Map<string, string>();
        for (const [attributeName, attributeValue] of rawAttributes) {
            if (attributeName === 'xmlns' || attributeName.startsWith('xmlns:')) continue;
            const attributeQName = splitQualifiedName(attributeName);
            if (attributeQName.prefix !== '' && !namespaces.has(attributeQName.prefix)) {
                return fail(`XML attribute prefix is not declared: ${attributeQName.prefix}`);
            }
            if (attributeQName.prefix !== '') {
                return fail('qualified attributes are not supported in OPC signature metadata');
            }
            if (attributes.has(attributeQName.localName)) return fail('XML contains ambiguous attributes');
            attributes.set(attributeQName.localName, attributeValue);
        }
        elements.push({ localName, namespaceUri, attributes, depth: stack.length });
        if (!selfClosing) stack.push({ qName, namespaces });
        else if (stack.length === 0) rootClosed = true;
        cursor = close + 1;
    }
    if (stack.length !== 0 || elements.length === 0 || !rootClosed) {
        return fail('XML document is incomplete');
    }
    return elements;
}

async function parseContentTypes(zip: ZipPackage): Promise<ContentTypeTable> {
    const entry = getEntry(zip, CONTENT_TYPES_PART);
    if (!entry) return fail('[Content_Types].xml is missing');
    const elements = parseXmlElements(await readEntry(zip, entry));
    const root = elements[0];
    if (
        root.localName !== 'Types' ||
        root.namespaceUri !== CONTENT_TYPES_NAMESPACE ||
        root.depth !== 0
    ) {
        return fail('[Content_Types].xml has an unexpected root element');
    }
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    for (const element of elements.slice(1)) {
        if (element.namespaceUri !== CONTENT_TYPES_NAMESPACE || element.depth !== 1) {
            return fail('[Content_Types].xml contains an unexpected namespace');
        }
        if (element.localName === 'Default') {
            const extension = element.attributes.get('Extension');
            const contentType = element.attributes.get('ContentType');
            if (!extension || !contentType || /[.\\/]/.test(extension)) {
                return fail('invalid Default content-type declaration');
            }
            const key = lowerAscii(extension);
            if (defaults.has(key)) return fail('duplicate Default content-type declaration');
            defaults.set(key, contentType);
        } else if (element.localName === 'Override') {
            const partName = element.attributes.get('PartName');
            const contentType = element.attributes.get('ContentType');
            if (!partName || !contentType) return fail('invalid Override content-type declaration');
            const key = lowerAscii(validatePartName(partName));
            if (overrides.has(key)) return fail('duplicate Override content-type declaration');
            overrides.set(key, contentType);
        } else {
            return fail(`unexpected element in [Content_Types].xml: ${element.localName}`);
        }
    }
    return { defaults, overrides };
}

function effectiveContentType(table: ContentTypeTable, partName: string): string | undefined {
    const override = table.overrides.get(lowerAscii(partName));
    if (override !== undefined) return override;
    const fileName = partName.slice(partName.lastIndexOf('/') + 1);
    const dot = fileName.lastIndexOf('.');
    return dot < 0 ? undefined : table.defaults.get(lowerAscii(fileName.slice(dot + 1)));
}

async function parseRelationships(zip: ZipPackage, partName: string): Promise<OpcRelationship[]> {
    const entry = getEntry(zip, partName);
    if (!entry) return fail(`relationship part is missing: ${partName}`);
    const elements = parseXmlElements(await readEntry(zip, entry));
    const root = elements[0];
    if (
        root.localName !== 'Relationships' ||
        root.namespaceUri !== RELATIONSHIPS_NAMESPACE ||
        root.depth !== 0
    ) {
        return fail(`relationship part has an unexpected root: ${partName}`);
    }
    const ids = new Set<string>();
    const relationships: OpcRelationship[] = [];
    for (const element of elements.slice(1)) {
        if (
            element.localName !== 'Relationship' ||
            element.namespaceUri !== RELATIONSHIPS_NAMESPACE ||
            element.depth !== 1
        ) {
            return fail(`unexpected element in relationship part: ${partName}`);
        }
        const id = element.attributes.get('Id');
        const type = element.attributes.get('Type');
        const target = element.attributes.get('Target');
        const targetMode = element.attributes.get('TargetMode');
        if (!id || !type || !target || ids.has(id)) {
            return fail(`invalid or duplicate relationship in ${partName}`);
        }
        ids.add(id);
        relationships.push({ id, type, target, targetMode });
    }
    return relationships;
}

function relationshipSourcePart(relationshipPartName: string): string {
    if (lowerAscii(relationshipPartName) === lowerAscii(ROOT_RELATIONSHIPS_PART)) return '/';
    const directory = posix.dirname(relationshipPartName);
    if (posix.basename(directory) !== '_rels') {
        return fail(`invalid relationship part location: ${relationshipPartName}`);
    }
    const parent = posix.dirname(directory);
    const fileName = posix.basename(relationshipPartName);
    if (!fileName.endsWith('.rels')) return fail('relationship part name is malformed');
    return validatePartName(posix.join(parent, fileName.slice(0, -5)));
}

function relationshipsPartFor(sourcePartName: string): string {
    validatePartName(sourcePartName);
    return validatePartName(
        posix.join(posix.dirname(sourcePartName), '_rels', `${posix.basename(sourcePartName)}.rels`),
    );
}

function resolveRelationshipTarget(sourcePartName: string, target: string): string {
    if (
        target.includes('\\') ||
        target.includes('\0') ||
        target.includes('?') ||
        target.includes('#') ||
        /^[A-Za-z][A-Za-z0-9+.-]*:/.test(target) ||
        target.startsWith('//')
    ) {
        return fail(`ambiguous or external OPC relationship target: ${target}`);
    }
    validatePercentEncoding(target);
    const baseDirectory = sourcePartName === '/' ? '/' : `${posix.dirname(sourcePartName)}/`;
    const unresolved = target.startsWith('/') ? target : `${baseDirectory}${target}`;
    const segments: string[] = [];
    for (const segment of unresolved.split('/')) {
        if (segment === '' && segments.length === 0) continue;
        if (segment === '' || segment === '.') continue;
        if (segment === '..') {
            if (segments.length === 0) return fail('relationship target escapes the package root');
            segments.pop();
        } else {
            segments.push(segment);
        }
    }
    return validatePartName(`/${segments.join('/')}`);
}

function isInternalRelationship(relationship: OpcRelationship): boolean {
    if (relationship.targetMode === undefined || relationship.targetMode === '') return true;
    if (relationship.targetMode === 'Internal') return true;
    return false;
}

async function inspectZipPackageForXlmMacroSheets(zip: ZipPackage): Promise<boolean> {
    const contentTypes = await parseContentTypes(zip);
    for (const overriddenPartName of contentTypes.overrides.keys()) {
        if (!zip.entries.has(overriddenPartName)) {
            return fail(`content-type override targets a missing part: ${overriddenPartName}`);
        }
    }

    let found = false;
    for (const entry of zip.entries.values()) {
        const partName = lowerAscii(entry.partName);
        if (
            partName.startsWith('/xl/macrosheets/') ||
            XLM_MACRO_SHEET_CONTENT_TYPES.has(effectiveContentType(contentTypes, entry.partName) ?? '')
        ) {
            found = true;
        }
    }

    const relationshipParts = [...zip.entries.values()].filter((entry) =>
        lowerAscii(entry.partName).endsWith('.rels') &&
        lowerAscii(posix.basename(posix.dirname(entry.partName))) === '_rels',
    );
    if (relationshipParts.length > MAX_RELATIONSHIP_PARTS) {
        return fail('package contains too many relationship parts');
    }
    let relationshipBytes = 0;
    let hasRootRelationships = false;
    for (const entry of relationshipParts) {
        relationshipBytes += entry.uncompressedSize;
        if (relationshipBytes > MAX_RELATIONSHIP_XML_BYTES) {
            return fail('relationship metadata exceeds the inspection limit');
        }
        if (lowerAscii(entry.partName) === lowerAscii(ROOT_RELATIONSHIPS_PART)) {
            hasRootRelationships = true;
        }
        const relationships = await parseRelationships(zip, entry.partName);
        if (relationships.some(relationship => XLM_MACRO_SHEET_RELATIONSHIP_TYPES.has(relationship.type))) {
            found = true;
        }
    }
    if (!hasRootRelationships) return fail('root relationship part is missing');
    return found;
}

async function inspectZipPackageSignature(zip: ZipPackage): Promise<boolean> {
    const contentTypes = await parseContentTypes(zip);
    for (const overriddenPartName of contentTypes.overrides.keys()) {
        if (!zip.entries.has(overriddenPartName)) {
            return fail(`content-type override targets a missing part: ${overriddenPartName}`);
        }
    }
    const relationshipParts = [...zip.entries.values()].filter((entry) =>
        lowerAscii(entry.partName).endsWith('.rels') &&
        lowerAscii(posix.basename(posix.dirname(entry.partName))) === '_rels',
    );
    if (relationshipParts.length > MAX_RELATIONSHIP_PARTS) {
        return fail('package contains too many relationship parts');
    }
    let relationshipBytes = 0;
    const parsedRelationships = new Map<string, OpcRelationship[]>();
    for (const entry of relationshipParts) {
        relationshipBytes += entry.uncompressedSize;
        if (relationshipBytes > MAX_RELATIONSHIP_XML_BYTES) {
            return fail('relationship metadata exceeds the inspection limit');
        }
        parsedRelationships.set(
            lowerAscii(entry.partName),
            await parseRelationships(zip, entry.partName),
        );
    }
    const rootRelationships = parsedRelationships.get(lowerAscii(ROOT_RELATIONSHIPS_PART));
    if (!rootRelationships) return fail('root relationship part is missing');

    const originTypedParts: string[] = [];
    const signatureTypedParts: string[] = [];
    for (const entry of zip.entries.values()) {
        const contentType = effectiveContentType(contentTypes, entry.partName);
        if (contentType === ORIGIN_CONTENT_TYPE) originTypedParts.push(entry.partName);
        if (contentType === SIGNATURE_CONTENT_TYPE) signatureTypedParts.push(entry.partName);
    }

    const rootOriginRelationships = rootRelationships.filter(
        (relationship) => relationship.type === ORIGIN_RELATIONSHIP_TYPE,
    );
    const relevantRelationships: Array<{
        relationshipPart: string;
        relationship: OpcRelationship;
    }> = [];
    for (const [relationshipPart, relationships] of parsedRelationships) {
        for (const relationship of relationships) {
            if (
                relationship.type === ORIGIN_RELATIONSHIP_TYPE ||
                relationship.type === SIGNATURE_RELATIONSHIP_TYPE
            ) {
                relevantRelationships.push({ relationshipPart, relationship });
            }
        }
    }

    if (rootOriginRelationships.length === 0) {
        if (
            originTypedParts.length !== 0 ||
            signatureTypedParts.length !== 0 ||
            relevantRelationships.length !== 0
        ) {
            return fail('orphaned digital-signature artifacts were found');
        }
        return false;
    }
    if (rootOriginRelationships.length !== 1) {
        return fail('package has multiple digital-signature origin relationships');
    }
    const originRelationship = rootOriginRelationships[0];
    if (!isInternalRelationship(originRelationship)) {
        return fail('digital-signature origin relationship is external');
    }
    const originPartName = resolveRelationshipTarget('/', originRelationship.target);
    const originEntry = getEntry(zip, originPartName);
    if (!originEntry) return fail('digital-signature origin target is missing');
    if (effectiveContentType(contentTypes, originPartName) !== ORIGIN_CONTENT_TYPE) {
        return fail('digital-signature origin content type is missing or inconsistent');
    }
    if ((await readEntry(zip, originEntry)).length !== 0) {
        return fail('digital-signature origin part must be empty');
    }
    if (
        originTypedParts.length !== 1 ||
        lowerAscii(originTypedParts[0]) !== lowerAscii(originPartName)
    ) {
        return fail('orphaned or ambiguous digital-signature origin parts were found');
    }
    if (
        relevantRelationships.some(
            ({ relationshipPart, relationship }) =>
                relationship.type === ORIGIN_RELATIONSHIP_TYPE &&
                lowerAscii(relationshipPart) !== lowerAscii(ROOT_RELATIONSHIPS_PART),
        )
    ) {
        return fail('digital-signature origin relationship appears in an unexpected part');
    }
    const allSignatureRelationships = relevantRelationships.filter(
        ({ relationship }) => relationship.type === SIGNATURE_RELATIONSHIP_TYPE,
    );
    const originRelationshipsPart = relationshipsPartFor(originPartName);
    const originRelationships = parsedRelationships.get(lowerAscii(originRelationshipsPart));
    if (!originRelationships) {
        if (signatureTypedParts.length !== 0 || allSignatureRelationships.length !== 0) {
            return fail('orphaned digital-signature artifacts were found');
        }
        return false;
    }
    const signatureRelationships = originRelationships.filter(
        (relationship) => relationship.type === SIGNATURE_RELATIONSHIP_TYPE,
    );
    if (signatureRelationships.length === 0) {
        if (signatureTypedParts.length !== 0 || allSignatureRelationships.length !== 0) {
            return fail('orphaned digital-signature artifacts were found');
        }
        return false;
    }
    const signatureTargets = new Set<string>();
    for (const relationship of signatureRelationships) {
        if (!isInternalRelationship(relationship)) {
            return fail('digital-signature relationship is external');
        }
        const targetPartName = resolveRelationshipTarget(originPartName, relationship.target);
        const targetEntry = getEntry(zip, targetPartName);
        if (!targetEntry) return fail('digital-signature target is missing');
        if (effectiveContentType(contentTypes, targetPartName) !== SIGNATURE_CONTENT_TYPE) {
            return fail('digital-signature content type is missing or inconsistent');
        }
        const key = lowerAscii(targetPartName);
        if (signatureTargets.has(key)) return fail('duplicate digital-signature targets were found');
        signatureTargets.add(key);
    }
    if (
        signatureTypedParts.length !== signatureTargets.size ||
        signatureTypedParts.some((partName) => !signatureTargets.has(lowerAscii(partName)))
    ) {
        return fail('orphaned or ambiguous digital-signature parts were found');
    }
    for (const { relationshipPart, relationship } of relevantRelationships) {
        const expectedPart = relationship.type === ORIGIN_RELATIONSHIP_TYPE
            ? ROOT_RELATIONSHIPS_PART
            : originRelationshipsPart;
        if (lowerAscii(relationshipPart) !== lowerAscii(expectedPart)) {
            return fail('digital-signature relationship appears in an unexpected part');
        }
    }
    return true;
}

export function isOoxmlPackagePath(fileName: string): boolean {
    return OOXML_SPREADSHEET_EXTENSIONS.has(extname(fileName).toLowerCase());
}

export async function hasOoxmlPackageSignatureBytes(
    bytes: Uint8Array,
    fileName = 'workbook.xlsx',
): Promise<boolean> {
    if (!isOoxmlPackagePath(fileName)) return false;
    return inspectZipPackageSignature(await openZipPackage(createBufferReader(bytes)));
}

export async function hasOoxmlPackageSignature(filePath: string): Promise<boolean> {
    if (!isOoxmlPackagePath(filePath)) return false;
    const handle = await open(filePath, 'r');
    try {
        const stats = await handle.stat();
        if (!stats.isFile()) return fail('source is not a regular file');
        return await inspectZipPackageSignature(
            await openZipPackage(createFileReader(handle, stats.size)),
        );
    } finally {
        await handle.close();
    }
}

export async function hasOoxmlXlmMacroSheetsBytes(
    bytes: Uint8Array,
    fileName = 'workbook.xlsx',
): Promise<boolean> {
    if (!isOoxmlPackagePath(fileName)) return false;
    return inspectZipPackageForXlmMacroSheets(await openZipPackage(createBufferReader(bytes)));
}

export async function hasOoxmlXlmMacroSheets(filePath: string): Promise<boolean> {
    if (!isOoxmlPackagePath(filePath)) return false;
    const handle = await open(filePath, 'r');
    try {
        const stats = await handle.stat();
        if (!stats.isFile()) return fail('source is not a regular file');
        return await inspectZipPackageForXlmMacroSheets(
            await openZipPackage(createFileReader(handle, stats.size)),
        );
    } finally {
        await handle.close();
    }
}

export async function assertOoxmlHasNoXlmMacroSheetsForAutomation(
    filePath: string,
): Promise<void> {
    try {
        if (await hasOoxmlXlmMacroSheets(filePath)) {
            throw new Error(OOXML_XLM_AUTOMATION_BLOCKED_MESSAGE);
        }
    } catch (error) {
        if (
            error instanceof Error &&
            error.message === OOXML_XLM_AUTOMATION_BLOCKED_MESSAGE
        ) {
            throw error;
        }
        throw new Error(OOXML_XLM_VERIFICATION_BLOCKED_MESSAGE, { cause: error });
    }
}

export async function assertOoxmlPackageUnsignedBytesForMutation(
    bytes: Uint8Array,
    fileName = 'workbook.xlsx',
): Promise<void> {
    try {
        if (await hasOoxmlPackageSignatureBytes(bytes, fileName)) {
            throw new Error(OOXML_PACKAGE_SIGNATURE_WRITE_BLOCKED_MESSAGE);
        }
    } catch (error) {
        if (
            error instanceof Error &&
            error.message === OOXML_PACKAGE_SIGNATURE_WRITE_BLOCKED_MESSAGE
        ) {
            throw error;
        }
        throw new Error(OOXML_PACKAGE_SIGNATURE_VERIFICATION_BLOCKED_MESSAGE, {
            cause: error,
        });
    }
}

export async function assertOoxmlPackageUnsignedForMutation(
    filePath: string,
): Promise<void> {
    try {
        if (await hasOoxmlPackageSignature(filePath)) {
            throw new Error(OOXML_PACKAGE_SIGNATURE_WRITE_BLOCKED_MESSAGE);
        }
    } catch (error) {
        if (
            error instanceof Error &&
            error.message === OOXML_PACKAGE_SIGNATURE_WRITE_BLOCKED_MESSAGE
        ) {
            throw error;
        }
        throw new Error(OOXML_PACKAGE_SIGNATURE_VERIFICATION_BLOCKED_MESSAGE, {
            cause: error,
        });
    }
}

export async function assertExistingOoxmlPackageUnsignedForMutation(
    filePath: string,
): Promise<void> {
    try {
        await assertOoxmlPackageUnsignedForMutation(filePath);
    } catch (error) {
        const code = (error as NodeJS.ErrnoException)?.cause
            ? ((error as Error & { cause?: NodeJS.ErrnoException }).cause?.code)
            : (error as NodeJS.ErrnoException)?.code;
        if (code === 'ENOENT') return;
        throw error;
    }
}
