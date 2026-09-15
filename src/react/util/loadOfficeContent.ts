export interface OfficeOpenPayload {
    path?: string;
    buffer?: number[];
    bufferBase64?: string;
    error?: string;
    ext?: string;
    documentCacheId?: string;
    readOnly?: boolean;
    readOnlyReason?: string;
    backupSheets?: unknown[];
    backupSourceSha256?: string;
    nativeLoadGeneration?: string;
}

export function parseOfficeOpenPayload(value: unknown): OfficeOpenPayload {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
        throw new Error('Invalid spreadsheet open message');
    }
    const payload = value as Record<string, unknown>;
    for (const key of ['path', 'bufferBase64', 'error', 'ext', 'documentCacheId', 'readOnlyReason', 'backupSourceSha256', 'nativeLoadGeneration']) {
        if (payload[key] !== undefined && typeof payload[key] !== 'string') {
            throw new Error(`Invalid spreadsheet open field: ${key}`);
        }
    }
    if (payload.readOnly !== undefined && typeof payload.readOnly !== 'boolean') {
        throw new Error('Invalid spreadsheet read-only flag');
    }
    if (payload.backupSheets !== undefined && !Array.isArray(payload.backupSheets)) {
        throw new Error('Invalid spreadsheet recovery data');
    }
    if (payload.buffer !== undefined && (!Array.isArray(payload.buffer) || payload.buffer.length > 135_000_000 || payload.buffer.some(byte => !Number.isInteger(byte) || byte < 0 || byte > 255))) {
        throw new Error('Invalid spreadsheet byte payload');
    }
    return payload as OfficeOpenPayload;
}

const MAX_BASE64_PAYLOAD_CHARACTERS = 180_000_000;

export function arrayBufferFromBase64Payload(value: string): ArrayBuffer {
    if (!value || value.length > MAX_BASE64_PAYLOAD_CHARACTERS || value.length % 4 !== 0) {
        throw new Error('Invalid or oversized base64 file content');
    }
    if (!/^[A-Za-z0-9+/]*={0,2}$/.test(value)) {
        throw new Error('Invalid base64 file content');
    }
    const padding = value.endsWith('==') ? 2 : value.endsWith('=') ? 1 : 0;
    const output = new Uint8Array((value.length / 4) * 3 - padding);
    const chunkCharacters = 1024 * 1024;
    let outputOffset = 0;
    for (let offset = 0; offset < value.length; offset += chunkCharacters) {
        const decoded = atob(value.slice(offset, offset + chunkCharacters));
        for (let index = 0; index < decoded.length; index += 1) {
            output[outputOffset++] = decoded.charCodeAt(index);
        }
    }
    if (outputOffset !== output.byteLength) throw new Error('Invalid base64 file length');
    return output.buffer;
}

export function arrayBufferFromPayload(payload: OfficeOpenPayload): ArrayBuffer {
    if (payload.error) {
        throw new Error(payload.error);
    }
    if (!payload.buffer?.length) {
        throw new Error('Empty file content');
    }
    const bytes = new Uint8Array(payload.buffer.length);
    for (let i = 0; i < payload.buffer.length; i++) {
        bytes[i] = payload.buffer[i];
    }
    return bytes.buffer;
}

export async function loadOfficeBuffer(payload: OfficeOpenPayload): Promise<ArrayBuffer> {
    if (payload.bufferBase64) {
        return arrayBufferFromBase64Payload(payload.bufferBase64);
    }
    if (payload.buffer) {
        return arrayBufferFromPayload(payload);
    }
    if (!payload.path) {
        throw new Error(payload.error ?? 'No file path');
    }
    const response = await fetch(payload.path);
    if (!response.ok) {
        throw new Error(`Failed to fetch (${response.status})`);
    }
    return response.arrayBuffer();
}
