import { Handler } from '@/common/handler';
import { isUriReadOnly } from '@/common/fileReadOnly';
import {
    assertExistingOoxmlPackageUnsignedForMutation,
    assertOoxmlPackageUnsignedBytesForMutation,
    assertOoxmlPackageUnsignedForMutation,
    hasOoxmlPackageSignature,
    hasOoxmlPackageSignatureBytes,
    isOoxmlPackagePath,
    MAX_VIRTUAL_OOXML_PACKAGE_BYTES,
} from '@/common/ooxmlPackageSignature';
import { promises as fs } from 'fs';
import { basename, extname } from 'path';
import { Uri, workspace } from 'vscode';

export type EmbeddedSpreadsheetReadOnlyReason =
    | 'macro-preservation'
    | 'native-excel-editing'
    | 'file-permissions'
    | 'package-signature'
    | 'package-signature-verification';

export interface OfficeOpenSnapshot {
    nativeLoadGeneration?: string;
    backupSheets?: unknown;
    backupSourceSha256?: string;
}

const MACRO_OR_LEGACY_EXTENSIONS = new Set(['.xlsm', '.xls']);
const MAX_LOCAL_OFFICE_OPEN_BYTES = 128 * 1024 * 1024;

/**
 * The embedded writer rebuilds a workbook and cannot guarantee preservation
 * of VBA projects or every legacy BIFF record. Keep those source formats
 * view-only and send edits to native Excel instead.
 */
export function requiresNativeExcelForEditing(uri: Uri): boolean {
    return MACRO_OR_LEGACY_EXTENSIONS.has(extname(uri.fsPath).toLowerCase());
}

/** XLSM files can be edited safely through targeted native Excel operations. */
export function supportsNativeMacroEditing(uri: Uri): boolean {
    return extname(uri.fsPath).toLowerCase() === '.xlsm';
}

function isFileNotFoundError(error: unknown): boolean {
    const candidate = error as { code?: string; name?: string } | undefined;
    return candidate?.code === 'FileNotFound'
        || candidate?.code === 'ENOENT'
        || candidate?.name === 'EntryNotFound (FileSystemError)';
}

async function readVirtualOoxmlBytesForInspection(uri: Uri): Promise<Uint8Array> {
    const stat = await workspace.fs.stat(uri);
    if (stat.size > MAX_VIRTUAL_OOXML_PACKAGE_BYTES) {
        throw new Error(
            `Virtual OOXML package exceeds the ${MAX_VIRTUAL_OOXML_PACKAGE_BYTES}-byte inspection limit.`,
        );
    }
    const bytes = await workspace.fs.readFile(uri);
    if (bytes.byteLength > MAX_VIRTUAL_OOXML_PACKAGE_BYTES) {
        throw new Error(
            `Virtual OOXML package exceeds the ${MAX_VIRTUAL_OOXML_PACKAGE_BYTES}-byte inspection limit.`,
        );
    }
    return bytes;
}

export async function hasUriOoxmlPackageSignature(uri: Uri): Promise<boolean> {
    if (!isOoxmlPackagePath(uri.fsPath)) return false;
    if (uri.scheme === 'file') {
        return hasOoxmlPackageSignature(uri.fsPath);
    }
    return hasOoxmlPackageSignatureBytes(
        await readVirtualOoxmlBytesForInspection(uri),
        uri.fsPath,
    );
}

/** Re-inspect the source bytes immediately before every host write. */
export async function assertUriOoxmlPackageUnsignedForMutation(uri: Uri): Promise<void> {
    if (!isOoxmlPackagePath(uri.fsPath)) return;
    if (uri.scheme === 'file') {
        await assertOoxmlPackageUnsignedForMutation(uri.fsPath);
        return;
    }
    await assertOoxmlPackageUnsignedBytesForMutation(
        await readVirtualOoxmlBytesForInspection(uri),
        uri.fsPath,
    );
}

/** Existing Save As destinations are protected for both file and virtual URIs. */
export async function assertExistingUriOoxmlPackageUnsignedForMutation(
    uri: Uri,
): Promise<void> {
    if (!isOoxmlPackagePath(uri.fsPath)) return;
    if (uri.scheme === 'file') {
        await assertExistingOoxmlPackageUnsignedForMutation(uri.fsPath);
        return;
    }
    try {
        await workspace.fs.stat(uri);
    } catch (error) {
        if (isFileNotFoundError(error)) return;
        throw error;
    }
    await assertUriOoxmlPackageUnsignedForMutation(uri);
}

export async function getEmbeddedSpreadsheetReadOnlyState(
    uri: Uri
): Promise<{
    readOnly: boolean;
    readOnlyReason?: EmbeddedSpreadsheetReadOnlyReason;
}> {
    if (isOoxmlPackagePath(uri.fsPath)) {
        try {
            if (await hasUriOoxmlPackageSignature(uri)) {
                return {
                    readOnly: true,
                    readOnlyReason: 'package-signature',
                };
            }
        } catch {
            return {
                readOnly: true,
                readOnlyReason: 'package-signature-verification',
            };
        }
    }
    if (await isUriReadOnly(uri)) {
        return {
            readOnly: true,
            readOnlyReason: 'file-permissions',
        };
    }
    if (supportsNativeMacroEditing(uri)) {
        return {
            readOnly: false,
            readOnlyReason: 'native-excel-editing',
        };
    }
    if (requiresNativeExcelForEditing(uri)) {
        return {
            readOnly: true,
            readOnlyReason: 'macro-preservation',
        };
    }
    return { readOnly: false };
}

export function isVirtualUri(uri: Uri): boolean {
    return uri.scheme !== 'file';
}

/** Stable per-document key used only for restoring the spreadsheet view state. */
export function buildDocumentCacheId(uri: Uri): string {
    return `${uri.scheme}:${uri.toString()}`;
}

async function readStableLocalFileBytes(
    filePath: string,
    maximumBytes: number,
    label: string,
    allowEmpty: boolean,
): Promise<Uint8Array> {
    const handle = await fs.open(filePath, 'r');
    try {
        const before = await handle.stat();
        if ((!allowEmpty && before.size <= 0) || before.size > maximumBytes) {
            throw new Error(`${label} has an invalid size (limit: ${maximumBytes} bytes).`);
        }
        const buffer = Buffer.allocUnsafe(before.size);
        let offset = 0;
        while (offset < buffer.byteLength) {
            const { bytesRead } = await handle.read(
                buffer,
                offset,
                buffer.byteLength - offset,
                offset,
            );
            if (bytesRead === 0) break;
            offset += bytesRead;
        }
        const after = await handle.stat();
        if (offset !== before.size || after.size !== before.size) {
            throw new Error(`${label} changed size while it was read.`);
        }
        return new Uint8Array(buffer.buffer, buffer.byteOffset, buffer.byteLength);
    } finally {
        await handle.close();
    }
}

export async function readUriBytesWithLimit(
    uri: Uri,
    maximumBytes: number,
    label: string,
    allowEmpty = true,
): Promise<Uint8Array> {
    if (!Number.isSafeInteger(maximumBytes) || maximumBytes < 1) {
        throw new Error('Office read limit is invalid.');
    }
    if (uri.scheme === 'file') {
        return readStableLocalFileBytes(uri.fsPath, maximumBytes, label, allowEmpty);
    }
    const stat = await workspace.fs.stat(uri);
    if ((!allowEmpty && stat.size <= 0) || stat.size > maximumBytes) {
        throw new Error(`${label} has an invalid size (limit: ${maximumBytes} bytes).`);
    }
    const bytes = await workspace.fs.readFile(uri);
    if ((!allowEmpty && bytes.byteLength <= 0) || bytes.byteLength > maximumBytes) {
        throw new Error(`${label} changed size while it was read.`);
    }
    return bytes;
}

export async function readUriBytes(uri: Uri): Promise<Uint8Array> {
    return readUriBytesWithLimit(
        uri,
        MAX_LOCAL_OFFICE_OPEN_BYTES,
        'Office file',
    );
}

export async function readUriText(uri: Uri): Promise<string> {
    return new TextDecoder('utf-8').decode(await readUriBytes(uri));
}

export function bytesToPayloadBuffer(data: Uint8Array): number[] {
    return Array.from(data);
}

/** Compact transport for virtual Office files; avoids expanding each byte to a JS number. */
export function bytesToPayloadBase64(data: Uint8Array): string {
    return Buffer.from(data.buffer, data.byteOffset, data.byteLength).toString('base64');
}

export async function emitVirtualOfficeOpen(
    handler: Handler,
    uri: Uri,
    snapshot?: OfficeOpenSnapshot
): Promise<void> {
    const ext = extname(uri.fsPath);
    const readOnlyState = await getEmbeddedSpreadsheetReadOnlyState(uri);
    const basePayload = {
        ext,
        path: uri.toString(),
        fileName: basename(uri.fsPath),
        scheme: uri.scheme,
        documentCacheId: buildDocumentCacheId(uri),
        ...readOnlyState,
        nonce: Date.now(),
    };

    try {
        const data = await readUriBytes(uri);
        handler.emit('open', {
            ...basePayload,
            ...snapshot,
            bufferBase64: bytesToPayloadBase64(data),
        });
    } catch (error) {
        handler.emit('open', {
            ...basePayload,
            ...snapshot,
            error: error instanceof Error ? error.message : 'Failed to read file',
        });
    }
}

export async function emitFileOfficeOpen(
    handler: Handler,
    uri: Uri,
    snapshot?: OfficeOpenSnapshot
): Promise<void> {
    const readOnlyState = await getEmbeddedSpreadsheetReadOnlyState(uri);
    const basePayload = {
        ext: extname(uri.fsPath),
        path: uri.toString(),
        fileName: basename(uri.fsPath),
        scheme: uri.scheme,
        documentCacheId: buildDocumentCacheId(uri),
        ...readOnlyState,
        nonce: Date.now(),
    };

    try {
        const data = await readUriBytes(uri);
        handler.emit('open', {
            ...basePayload,
            ...snapshot,
            bufferBase64: bytesToPayloadBase64(data),
        });
    } catch (error) {
        handler.emit('open', {
            ...basePayload,
            ...snapshot,
            error: error instanceof Error ? error.message : 'Failed to read file',
        });
    }
}
