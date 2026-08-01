/** 主流 AI HTTP API 格式 */
export type CustomAIApiFormat = "auto" | "openai" | "anthropic" | "gemini" | "ollama";

export interface CustomAIRequestOptions {
    url: string;
    apiKey?: string;
    model?: string;
    format?: CustomAIApiFormat;
    prompt: string;
    signal?: AbortSignal;
}

const DEFAULT_MODELS: Record<Exclude<CustomAIApiFormat, "auto">, string> = {
    openai: "gpt-4o",
    anthropic: "claude-3-5-sonnet-20241022",
    gemini: "gemini-1.5-flash",
    ollama: "llama3.2",
};

const MAX_AI_RESPONSE_BYTES = 16 * 1024 * 1024;
const MAX_AI_RESPONSE_CHUNKS = 16 * 1024;
const MAX_AI_STREAM_RESPONSE_BYTES = 64 * 1024 * 1024;
const MAX_AI_STREAM_BUFFER_CHARACTERS = 1024 * 1024;

const parseApiUrl = (rawUrl: string): URL | undefined => {
    try {
        return new URL(rawUrl);
    } catch {
        return undefined;
    }
};

const normalizeHostname = (hostname: string): string => hostname.toLowerCase().replace(/\.$/, "");

const hostnameMatchesDomain = (hostname: string, domain: string): boolean => {
    const normalizedHostname = normalizeHostname(hostname);
    return normalizedHostname === domain || normalizedHostname.endsWith(`.${domain}`);
};

const isLoopbackHostname = (hostname: string): boolean => {
    const normalizedHostname = normalizeHostname(hostname);
    return normalizedHostname === "localhost"
        || normalizedHostname === "[::1]"
        || /^127(?:\.\d{1,3}){3}$/.test(normalizedHostname);
};

const hasQueryApiKey = (url: URL): boolean => {
    for (const key of url.searchParams.keys()) {
        if (key.toLowerCase() === "key") {
            return true;
        }
    }
    return false;
};

const validateApiUrl = (rawUrl: string): string => {
    const parsed = parseApiUrl(rawUrl);
    if (!parsed || (parsed.protocol !== "https:" && parsed.protocol !== "http:")) {
        throw new Error("API URL must be an absolute HTTP(S) URL.");
    }
    if (parsed.username || parsed.password) {
        throw new Error("API URL must not contain embedded credentials.");
    }
    if (hasQueryApiKey(parsed)) {
        throw new Error("API URL must not contain a 'key' query parameter. Use the API key field instead.");
    }
    if (parsed.protocol === "http:" && !isLoopbackHostname(parsed.hostname)) {
        throw new Error("HTTPS is required for non-local AI endpoints.");
    }
    return parsed.toString();
};

const trimTrailingPathSlash = (pathname: string): string => pathname.replace(/\/+$/, "");

const withPathname = (url: URL, pathname: string): string => {
    url.pathname = pathname || "/";
    return url.toString();
};

/** 根据 URL 猜测 API 格式 */
export const detectCustomAIFormat = (url: string): Exclude<CustomAIApiFormat, "auto"> => {
    const parsed = parseApiUrl(url);
    if (!parsed) {
        return "openai";
    }

    const pathname = parsed.pathname.toLowerCase();
    if (hostnameMatchesDomain(parsed.hostname, "anthropic.com") || pathname.endsWith("/v1/messages")) {
        return "anthropic";
    }
    if (
        hostnameMatchesDomain(parsed.hostname, "generativelanguage.googleapis.com")
        || hostnameMatchesDomain(parsed.hostname, "googleapis.com")
        || pathname.endsWith(":generatecontent")
        || pathname.split("/").some((segment) => segment.startsWith("gemini"))
    ) {
        return "gemini";
    }
    if (
        parsed.port === "11434"
        || pathname.endsWith("/api/chat")
    ) {
        return "ollama";
    }
    return "openai";
};

const resolveFormat = (url: string, format?: CustomAIApiFormat): Exclude<CustomAIApiFormat, "auto"> => {
    if (format && format !== "auto") {
        return format;
    }
    return detectCustomAIFormat(url);
};

const resolveModel = (format: Exclude<CustomAIApiFormat, "auto">, model?: string) => {
    return model?.trim() || DEFAULT_MODELS[format];
};

const resolveOpenAIUrl = (rawUrl: string) => {
    const url = new URL(rawUrl);
    const pathname = trimTrailingPathSlash(url.pathname);
    if (pathname.endsWith("/chat/completions")) {
        return withPathname(url, pathname);
    }
    if (pathname.endsWith("/v1")) {
        return withPathname(url, `${pathname}/chat/completions`);
    }
    if (/\/v\d+$/.test(pathname)) {
        return withPathname(url, `${pathname}/chat/completions`);
    }
    return withPathname(url, `${pathname}/v1/chat/completions`);
};

const resolveAnthropicUrl = (rawUrl: string) => {
    const url = new URL(rawUrl);
    const pathname = trimTrailingPathSlash(url.pathname);
    if (pathname.endsWith("/messages")) {
        return withPathname(url, pathname);
    }
    if (pathname.endsWith("/v1")) {
        return withPathname(url, `${pathname}/messages`);
    }
    return withPathname(url, `${pathname}/v1/messages`);
};

const resolveOllamaUrl = (rawUrl: string) => {
    const url = new URL(rawUrl);
    const pathname = trimTrailingPathSlash(url.pathname);
    if (pathname.endsWith("/api/chat")) {
        return withPathname(url, pathname);
    }
    if (pathname.endsWith("/api")) {
        return withPathname(url, `${pathname}/chat`);
    }
    return withPathname(url, `${pathname}/api/chat`);
};

const resolveGeminiUrl = (rawUrl: string, model: string) => {
    const url = new URL(rawUrl);
    const pathname = trimTrailingPathSlash(url.pathname);
    const normalizedPathname = pathname.toLowerCase();
    if (normalizedPathname.endsWith(":generatecontent")) {
        return withPathname(url, pathname);
    }
    if (pathname.split("/").includes("models")) {
        return withPathname(url, `${pathname}:generateContent`);
    }
    const basePath = hostnameMatchesDomain(url.hostname, "googleapis.com")
        ? pathname
        : `${pathname}/v1beta`;
    return withPathname(url, `${basePath}/models/${encodeURIComponent(model)}:generateContent`);
};

const buildRequest = (
    format: Exclude<CustomAIApiFormat, "auto">,
    prompt: string,
    model: string,
    rawUrl: string,
    apiKey?: string,
) => {
    switch (format) {
        case "anthropic":
            return {
                url: resolveAnthropicUrl(rawUrl),
                headers: {
                    "Content-Type": "application/json",
                    "anthropic-version": "2023-06-01",
                    ...(apiKey ? { "x-api-key": apiKey } : {}),
                },
                body: {
                    model,
                    max_tokens: 8192,
                    messages: [{ role: "user", content: prompt }],
                },
            };
        case "gemini":
            const geminiUrl = resolveGeminiUrl(rawUrl, model);
            return {
                url: geminiUrl,
                headers: {
                    "Content-Type": "application/json",
                    ...(apiKey ? { "x-goog-api-key": apiKey } : {}),
                },
                body: {
                    contents: [{ role: "user", parts: [{ text: prompt }] }],
                },
            };
        case "ollama":
            return {
                url: resolveOllamaUrl(rawUrl),
                headers: {
                    "Content-Type": "application/json",
                    ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {}),
                },
                body: {
                    model,
                    messages: [{ role: "user", content: prompt }],
                    stream: false,
                },
            };
        default:
            return {
                url: resolveOpenAIUrl(rawUrl),
                headers: {
                    "Content-Type": "application/json",
                    ...(apiKey ? { Authorization: `Bearer ${apiKey}` } : {}),
                },
                body: {
                    model,
                    messages: [{ role: "user", content: prompt }],
                },
            };
    }
};

const readTextContent = (value: unknown): string => {
    if (typeof value === "string") {
        return value;
    }
    if (!value || typeof value !== "object") {
        return "";
    }
    const record = value as Record<string, unknown>;
    if (typeof record.text === "string") {
        return record.text;
    }
    if (Array.isArray(record.content)) {
        let text = "";
        for (const item of record.content) {
            text += readTextContent(item);
        }
        return text;
    }
    if (Array.isArray(record.parts)) {
        let text = "";
        for (const part of record.parts) {
            text += readTextContent(part);
        }
        return text;
    }
    return "";
};

const parseOpenAIResponse = (data: any): string => {
    const choice = data?.choices?.[0];
    const messageContent = choice?.message?.content;
    if (typeof messageContent === "string") {
        return messageContent;
    }
    if (Array.isArray(messageContent)) {
        let text = "";
        for (const part of messageContent) {
            if (part?.type === "text" && typeof part.text === "string") {
                text += part.text;
            }
        }
        if (text) {
            return text;
        }
    }
    if (typeof choice?.text === "string") {
        return choice.text;
    }
    if (typeof choice?.message?.text === "string") {
        return choice.message.text;
    }
    return "";
};

const parseAnthropicResponse = (data: any): string => {
    if (!Array.isArray(data?.content)) {
        return "";
    }
    let text = "";
    for (const block of data.content) {
        if (block?.type === "text" && typeof block.text === "string") {
            text += block.text;
        }
    }
    return text;
};

const parseGeminiResponse = (data: any): string => {
    const parts = data?.candidates?.[0]?.content?.parts;
    if (!Array.isArray(parts)) {
        return "";
    }
    let text = "";
    for (const part of parts) {
        if (typeof part?.text === "string") {
            text += part.text;
        }
    }
    return text;
};

const parseOllamaResponse = (data: any): string => {
    if (typeof data?.message?.content === "string") {
        return data.message.content;
    }
    if (typeof data?.response === "string") {
        return data.response;
    }
    return "";
};

const parseGenericResponse = (data: any): string => {
    const direct = [
        data?.output,
        data?.response,
        data?.result,
        data?.text,
        data?.content,
        data?.data?.output,
        data?.data?.text,
    ];
    for (const item of direct) {
        const text = readTextContent(item);
        if (text.trim()) {
            return text;
        }
    }
    return "";
};

const parseResponseText = (
    format: Exclude<CustomAIApiFormat, "auto">,
    data: any,
): string => {
    const parsers: Array<() => string> = [];
    switch (format) {
        case "anthropic":
            parsers.push(() => parseAnthropicResponse(data));
            break;
        case "gemini":
            parsers.push(() => parseGeminiResponse(data));
            break;
        case "ollama":
            parsers.push(() => parseOllamaResponse(data));
            break;
        default:
            parsers.push(() => parseOpenAIResponse(data));
            break;
    }
    parsers.push(
        () => parseOpenAIResponse(data),
        () => parseAnthropicResponse(data),
        () => parseGeminiResponse(data),
        () => parseOllamaResponse(data),
        () => parseGenericResponse(data),
    );

    for (const parser of parsers) {
        const text = parser().trim();
        if (text) {
            return text;
        }
    }
    return "";
};

const parseErrorMessage = (data: any, status: number, statusText: string): string => {
    const message = data?.error?.message
        ?? data?.error?.msg
        ?? data?.message
        ?? data?.detail
        ?? (typeof data?.error === "string" ? data.error : undefined);
    if (typeof message === "string" && message.trim()) {
        return message;
    }
    return `HTTP ${status} ${statusText}`;
};

const declaredContentLength = (response: Response): number | undefined => {
    const rawLength = response.headers.get("content-length")?.trim();
    if (!rawLength || !/^\d+$/.test(rawLength)) {
        return undefined;
    }
    const length = Number(rawLength);
    return Number.isSafeInteger(length) ? length : Number.POSITIVE_INFINITY;
};

const cancelReader = async (reader: ReadableStreamDefaultReader<Uint8Array>): Promise<void> => {
    try {
        await reader.cancel();
    } catch {
        // Preserve the deterministic size-limit error if cancellation itself fails.
    }
};

class ResponseReadLimitError extends Error {
    constructor(message: string) {
        super(message);
        this.name = "ResponseReadLimitError";
    }
}

const responseTooLargeError = (limit: number): Error =>
    new ResponseReadLimitError(`AI API response exceeds the ${limit}-byte limit.`);

const responseChunkLimitError = (): Error =>
    new ResponseReadLimitError(`AI API response exceeds the ${MAX_AI_RESPONSE_CHUNKS}-chunk limit.`);

const readResponseTextLimited = async (
    response: Response,
    limit = MAX_AI_RESPONSE_BYTES,
): Promise<string> => {
    const reader = response.body?.getReader();
    const contentLength = declaredContentLength(response);
    if (contentLength !== undefined && contentLength > limit) {
        if (reader) await cancelReader(reader);
        throw responseTooLargeError(limit);
    }
    if (!reader) return "";

    const chunks: Uint8Array[] = [];
    let totalBytes = 0;
    let chunkCount = 0;
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        chunkCount += 1;
        if (chunkCount > MAX_AI_RESPONSE_CHUNKS) {
            await cancelReader(reader);
            throw responseChunkLimitError();
        }
        if (totalBytes + value.byteLength > limit) {
            await cancelReader(reader);
            throw responseTooLargeError(limit);
        }
        if (value.byteLength > 0) {
            chunks.push(value);
        }
        totalBytes += value.byteLength;
    }

    const bytes = new Uint8Array(totalBytes);
    let offset = 0;
    for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.byteLength;
    }
    return new TextDecoder().decode(bytes);
};

export interface StreamCustomAIOptions extends CustomAIRequestOptions {
    onChunk: (chunk: string) => void;
}

/** 从 SSE 行中提取 delta 文本，兼容 OpenAI / Anthropic / Ollama / Gemini 流格式 */
const parseStreamChunk = (format: Exclude<CustomAIApiFormat, "auto">, line: string): string => {
    const dataPrefix = "data: ";
    if (line.startsWith(dataPrefix)) {
        const raw = line.slice(dataPrefix.length).trim();
        if (raw === "[DONE]") return "";
        try {
            const obj = JSON.parse(raw);
            switch (format) {
                case "anthropic":
                    return obj?.delta?.text ?? "";
                case "gemini":
                    return obj?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
                default:
                    return obj?.choices?.[0]?.delta?.content ?? "";
            }
        } catch { return ""; }
    }
    // Ollama: newline-delimited JSON (no "data:" prefix)
    if (format === "ollama" && line.trim()) {
        try {
            const obj = JSON.parse(line);
            return obj?.message?.content ?? obj?.response ?? "";
        } catch { return ""; }
    }
    return "";
};

/** 流式请求 body（在 non-stream body 上加 stream: true） */
const buildStreamBody = (
    format: Exclude<CustomAIApiFormat, "auto">,
    body: Record<string, unknown>,
): Record<string, unknown> => {
    switch (format) {
        case "anthropic":
            return { ...body, stream: true };
        case "gemini":
            // Gemini 流式需要换用 streamGenerateContent 端点，URL 已在 resolveGeminiUrl 中处理
            return body;
        case "ollama":
            return { ...body, stream: true };
        default:
            return { ...body, stream: true };
    }
};

/** 将 Gemini URL 从 generateContent 换为 streamGenerateContent */
const toGeminiStreamUrl = (url: string): string =>
    url.replace(/:generateContent(\?|$)/, ":streamGenerateContent$1");

/** 流式调用自定义 AI，每个 token 通过 onChunk 回调，全部结束后 resolve */
export const streamCustomAI = async (options: StreamCustomAIOptions): Promise<void> => {
    const rawUrl = options.url?.trim();
    if (!rawUrl) throw new Error("API URL is required.");

    const validatedUrl = validateApiUrl(rawUrl);
    const format = resolveFormat(validatedUrl, options.format);
    const model = resolveModel(format, options.model);
    const request = buildRequest(format, options.prompt, model, validatedUrl, options.apiKey?.trim());

    const streamUrl = format === "gemini" ? toGeminiStreamUrl(request.url) : request.url;
    const streamBody = buildStreamBody(format, request.body as Record<string, unknown>);

    const resp = await fetch(streamUrl, {
        method: "POST",
        headers: request.headers,
        body: JSON.stringify(streamBody),
        signal: options.signal,
        redirect: "error",
    });

    if (!resp.ok) {
        const rawText = await readResponseTextLimited(resp).catch((error) => {
            if (error instanceof ResponseReadLimitError) throw error;
            return "";
        });
        let data: any;
        try { data = rawText ? JSON.parse(rawText) : {}; } catch { data = {}; }
        throw new Error(parseErrorMessage(data, resp.status, resp.statusText));
    }

    const reader = resp.body?.getReader();
    if (!reader) throw new Error("No response body.");
    const contentLength = declaredContentLength(resp);
    if (contentLength !== undefined && contentLength > MAX_AI_STREAM_RESPONSE_BYTES) {
        await cancelReader(reader);
        throw responseTooLargeError(MAX_AI_STREAM_RESPONSE_BYTES);
    }

    const decoder = new TextDecoder();
    let buf = "";
    let totalBytes = 0;
    let newlineSearchStart = 0;
    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        if (totalBytes + value.byteLength > MAX_AI_STREAM_RESPONSE_BYTES) {
            await cancelReader(reader);
            throw responseTooLargeError(MAX_AI_STREAM_RESPONSE_BYTES);
        }
        totalBytes += value.byteLength;
        buf += decoder.decode(value, { stream: true });
        let lineStart = 0;
        let newlineIndex = buf.indexOf("\n", newlineSearchStart);
        while (newlineIndex !== -1) {
            const line = buf.slice(lineStart, newlineIndex);
            const chunk = parseStreamChunk(format, line);
            if (chunk) options.onChunk(chunk);
            lineStart = newlineIndex + 1;
            newlineIndex = buf.indexOf("\n", lineStart);
        }
        if (lineStart > 0) {
            buf = buf.slice(lineStart);
        }
        newlineSearchStart = buf.length;
        if (buf.length > MAX_AI_STREAM_BUFFER_CHARACTERS) {
            await cancelReader(reader);
            throw new Error(
                `AI API streaming buffer exceeds the ${MAX_AI_STREAM_BUFFER_CHARACTERS}-character limit without a line break.`,
            );
        }
    }
    buf += decoder.decode();
    if (buf.length > MAX_AI_STREAM_BUFFER_CHARACTERS) {
        throw new Error(
            `AI API streaming buffer exceeds the ${MAX_AI_STREAM_BUFFER_CHARACTERS}-character limit without a line break.`,
        );
    }
    // flush remaining buffer
    if (buf.trim()) {
        const chunk = parseStreamChunk(format, buf);
        if (chunk) options.onChunk(chunk);
    }
};

/** 调用自定义 AI HTTP 接口，自动适配 OpenAI / Anthropic / Gemini / Ollama 等格式 */
export const callCustomAI = async (options: CustomAIRequestOptions): Promise<string> => {
    const rawUrl = options.url?.trim();
    if (!rawUrl) {
        throw new Error("API URL is required.");
    }

    const validatedUrl = validateApiUrl(rawUrl);
    const format = resolveFormat(validatedUrl, options.format);
    const model = resolveModel(format, options.model);
    const request = buildRequest(format, options.prompt, model, validatedUrl, options.apiKey?.trim());

    const resp = await fetch(request.url, {
        method: "POST",
        headers: request.headers,
        body: JSON.stringify(request.body),
        signal: options.signal,
        redirect: "error",
    });

    const rawText = await readResponseTextLimited(resp);
    let data: any;
    try {
        data = rawText ? JSON.parse(rawText) : {};
    } catch {
        if (!resp.ok) {
            throw new Error(`HTTP ${resp.status} ${resp.statusText}`);
        }
        return rawText.trim();
    }

    if (!resp.ok) {
        throw new Error(parseErrorMessage(data, resp.status, resp.statusText));
    }

    const result = parseResponseText(format, data);
    if (!result) {
        throw new Error("Empty response from AI API.");
    }
    return result;
};
