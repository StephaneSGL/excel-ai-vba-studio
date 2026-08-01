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
import { basename, extname } from 'path';
import { Uri, workspace, type Webview } from 'vscode';

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

export async function readUriBytes(uri: Uri): Promise<Uint8Array> {
    return workspace.fs.readFile(uri);
}

export async function readUriText(uri: Uri): Promise<string> {
    return new TextDecoder('utf-8').decode(await readUriBytes(uri));
}

export function bytesToPayloadBuffer(data: Uint8Array): number[] {
    return Array.from(data);
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
            buffer: bytesToPayloadBuffer(data),
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
    webview: Webview,
    snapshot?: OfficeOpenSnapshot
): Promise<void> {
    const readOnlyState = await getEmbeddedSpreadsheetReadOnlyState(uri);
    handler.emit('open', {
        ext: extname(uri.fsPath),
        path: webview.asWebviewUri(uri).toString(),
        fileName: basename(uri.fsPath),
        documentCacheId: buildDocumentCacheId(uri),
        ...readOnlyState,
        ...snapshot,
    });
}
