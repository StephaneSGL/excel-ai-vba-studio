export interface OfficeOpenPayload {
    path?: string;
    buffer?: number[];
    bufferBase64?: string;
    error?: string;
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
