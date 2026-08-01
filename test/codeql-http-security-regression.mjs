import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { build } from "esbuild";

const repositoryRoot = path.resolve(import.meta.dirname, "..");
const customAIClientPath = path.join(repositoryRoot, "src", "service", "ai", "customAIClient.ts");
const httpClientPath = path.join(repositoryRoot, "src", "provider", "http", "utils", "httpClient.ts");
const MAX_AI_RESPONSE_BYTES = 16 * 1024 * 1024;
const MAX_AI_STREAM_BUFFER_CHARACTERS = 1024 * 1024;
const MAX_HTTP_RESPONSE_BYTES = 64 * 1024 * 1024;
const MAX_RESPONSE_CHUNKS = 16 * 1024;

const customAIClientSource = await fs.readFile(customAIClientPath, "utf8");
const httpClientSource = await fs.readFile(httpClientPath, "utf8");
assert.match(customAIClientSource, /MAX_AI_RESPONSE_CHUNKS/);
assert.match(customAIClientSource, /newlineSearchStart/);
assert.match(customAIClientSource, /indexOf\("\\n", newlineSearchStart\)/);
assert.match(customAIClientSource, /newlineSearchStart\s*=\s*buf\.length/);
assert.doesNotMatch(httpClientSource, /rejectUnauthorized\s*:\s*false/);
assert.doesNotMatch(httpClientSource, /new\s+Agent\s*\(/);
assert.match(httpClientSource, /MAX_HTTP_RESPONSE_BYTES/);
assert.match(httpClientSource, /MAX_HTTP_RESPONSE_CHUNKS/);
assert.match(httpClientSource, /reader\.cancel\(\)/);

const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "excel-ai-url-security-"));
const compiledModulePath = path.join(temporaryDirectory, "custom-ai-client.mjs");
const compiledHttpClientPath = path.join(temporaryDirectory, "http-client.mjs");

try {
    await build({
        entryPoints: [customAIClientPath],
        outfile: compiledModulePath,
        bundle: true,
        format: "esm",
        platform: "node",
        target: "node20",
        logLevel: "silent",
    });

    const { callCustomAI, detectCustomAIFormat, streamCustomAI } =
        await import(pathToFileURL(compiledModulePath).href);

    assert.equal(detectCustomAIFormat("https://api.anthropic.com"), "anthropic");
    assert.equal(detectCustomAIFormat("https://proxy.example.test/v1/messages"), "anthropic");
    assert.equal(detectCustomAIFormat("https://anthropic.com.attacker.test"), "openai");
    assert.equal(detectCustomAIFormat("https://generativelanguage.googleapis.com"), "gemini");
    assert.equal(detectCustomAIFormat("https://edge.googleapis.com"), "gemini");
    assert.equal(detectCustomAIFormat("https://googleapis.com.attacker.test"), "openai");
    assert.equal(detectCustomAIFormat("http://localhost:11434"), "ollama");
    assert.equal(detectCustomAIFormat("https://proxy.example.test/api/chat"), "ollama");
    assert.equal(detectCustomAIFormat("not a URL"), "openai");

    const requests = [];
    const originalHttpFetch = globalThis.fetch;
    globalThis.fetch = async (url, options) => {
        requests.push({ url: String(url), options });
        return new Response(
            JSON.stringify({ choices: [{ message: { content: "ok" } }] }),
            {
                status: 200,
                statusText: "OK",
            },
        );
    };

    try {
        await assert.rejects(
            callCustomAI({
                url: "http://api.example.test",
                apiKey: "must-not-be-sent",
                prompt: "test",
            }),
            /HTTPS is required/,
        );
        await assert.rejects(
            callCustomAI({
                url: "https://user:password@api.example.test",
                prompt: "test",
            }),
            /must not contain embedded credentials/,
        );
        await assert.rejects(
            callCustomAI({
                url: "https://generativelanguage.googleapis.com/v1beta?KEY=legacy-secret",
                apiKey: "replacement-secret",
                prompt: "test",
            }),
            /must not contain a 'key' query parameter/,
        );
        await assert.rejects(
            callCustomAI({
                url: "https://generativelanguage.googleapis.com/v1beta?%6B%65%79=encoded-secret",
                prompt: "test",
            }),
            /must not contain a 'key' query parameter/,
        );
        assert.equal(requests.length, 0, "invalid endpoints must be rejected before fetch");

        await callCustomAI({
            url: "https://anthropic.com.attacker.test/base?tenant=one",
            apiKey: "openai-secret",
            prompt: "test",
        });
        assert.equal(requests[0].url, "https://anthropic.com.attacker.test/base/v1/chat/completions?tenant=one");
        assert.equal(requests[0].options.headers.Authorization, "Bearer openai-secret");
        assert.equal(requests[0].options.headers["x-api-key"], undefined);
        assert.equal(requests[0].options.redirect, "error");

        await callCustomAI({
            url: "https://generativelanguage.googleapis.com/v1beta",
            apiKey: "gemini-secret",
            model: "gemini/test model",
            prompt: "test",
        });
        const geminiUrl = new URL(requests[1].url);
        assert.equal(geminiUrl.hostname, "generativelanguage.googleapis.com");
        assert.equal(geminiUrl.searchParams.has("key"), false, "API keys must not be added to URLs");
        assert.equal(requests[1].options.headers["x-goog-api-key"], "gemini-secret");
        assert.match(geminiUrl.pathname, /\/models\/gemini%2Ftest%20model:generateContent$/);
        assert.equal(requests[1].options.redirect, "error");

        await callCustomAI({
            url: "http://127.0.0.1:11434",
            format: "ollama",
            prompt: "test",
        });
        assert.equal(requests[2].url, "http://127.0.0.1:11434/api/chat");

        let oversizedResponseCancelled = false;
        globalThis.fetch = async () => new Response(
            new ReadableStream({
                start(controller) {
                    controller.enqueue(new Uint8Array([1]));
                },
                cancel() {
                    oversizedResponseCancelled = true;
                },
            }),
            {
                status: 200,
                headers: { "content-length": String(MAX_AI_RESPONSE_BYTES + 1) },
            },
        );
        await assert.rejects(
            callCustomAI({
                url: "https://api.example.test/v1",
                prompt: "test",
            }),
            new RegExp(`exceeds the ${MAX_AI_RESPONSE_BYTES}-byte limit`),
        );
        assert.equal(oversizedResponseCancelled, true, "oversized response reader must be cancelled");

        let fragmentedResponseCancelled = false;
        let responseChunksProduced = 0;
        globalThis.fetch = async () => new Response(
            new ReadableStream({
                pull(controller) {
                    responseChunksProduced += 1;
                    controller.enqueue(new Uint8Array([0x20]));
                },
                cancel() {
                    fragmentedResponseCancelled = true;
                },
            }),
            { status: 200 },
        );
        await assert.rejects(
            callCustomAI({
                url: "https://api.example.test/v1",
                prompt: "test",
            }),
            new RegExp(`exceeds the ${MAX_RESPONSE_CHUNKS}-chunk limit`),
        );
        assert.ok(
            responseChunksProduced >= MAX_RESPONSE_CHUNKS + 1
            && responseChunksProduced <= MAX_RESPONSE_CHUNKS + 2,
            `unexpected response chunk prefetch count: ${responseChunksProduced}`,
        );
        assert.equal(fragmentedResponseCancelled, true, "over-fragmented response reader must be cancelled");

        let oversizedStreamCancelled = false;
        globalThis.fetch = async () => new Response(
            new ReadableStream({
                start(controller) {
                    controller.enqueue(
                        new TextEncoder().encode("x".repeat(MAX_AI_STREAM_BUFFER_CHARACTERS + 1)),
                    );
                },
                cancel() {
                    oversizedStreamCancelled = true;
                },
            }),
            { status: 200 },
        );
        await assert.rejects(
            streamCustomAI({
                url: "https://api.example.test/v1",
                prompt: "test",
                onChunk() {},
            }),
            new RegExp(`exceeds the ${MAX_AI_STREAM_BUFFER_CHARACTERS}-character limit`),
        );
        assert.equal(oversizedStreamCancelled, true, "oversized streaming reader must be cancelled");

        const fragmentedStreamChunkCount = 30_000;
        let streamChunksProduced = 0;
        globalThis.fetch = async () => new Response(
            new ReadableStream({
                pull(controller) {
                    if (streamChunksProduced < fragmentedStreamChunkCount) {
                        streamChunksProduced += 1;
                        controller.enqueue(new Uint8Array([0x78]));
                    } else {
                        controller.close();
                    }
                },
            }),
            { status: 200 },
        );
        const fragmentedStreamStart = performance.now();
        await streamCustomAI({
            url: "https://api.example.test/v1",
            prompt: "test",
            onChunk() {},
        });
        const fragmentedStreamDuration = performance.now() - fragmentedStreamStart;
        assert.equal(streamChunksProduced, fragmentedStreamChunkCount);
        assert.ok(
            fragmentedStreamDuration < 4_000,
            `fragmented stream cursor regression took ${Math.round(fragmentedStreamDuration)} ms`,
        );
    } finally {
        globalThis.fetch = originalHttpFetch;
    }

    await build({
        stdin: {
            contents: [
                "export { HttpClient } from './src/provider/http/utils/httpClient.ts';",
                "export { HttpRequest } from './src/provider/http/models/httpRequest.ts';",
            ].join("\n"),
            resolveDir: repositoryRoot,
            sourcefile: "http-client-security-entry.ts",
            loader: "ts",
        },
        outfile: compiledHttpClientPath,
        bundle: true,
        format: "esm",
        platform: "node",
        target: "node20",
        logLevel: "silent",
        plugins: [{
            name: "vscode-test-stub",
            setup(buildApi) {
                buildApi.onResolve({ filter: /^vscode$/ }, () => ({
                    path: "vscode",
                    namespace: "vscode-test-stub",
                }));
                buildApi.onLoad({ filter: /.*/, namespace: "vscode-test-stub" }, () => ({
                    loader: "js",
                    contents: `
                        export class EventEmitter {
                            constructor() { this.event = () => ({ dispose() {} }); }
                            fire() {}
                        }
                        export const languages = { setLanguageConfiguration() {} };
                        export const ViewColumn = { Active: 1, Beside: 2 };
                        export const window = {
                            activeTextEditor: undefined,
                            onDidChangeActiveTextEditor() { return { dispose() {} }; },
                        };
                        export const workspace = {
                            getConfiguration() { return { get(_key, fallback) { return fallback; } }; },
                            getWorkspaceFolder() { return undefined; },
                            onDidChangeConfiguration() { return { dispose() {} }; },
                        };
                    `,
                }));
            },
        }],
    });

    const { HttpClient, HttpRequest } = await import(pathToFileURL(compiledHttpClientPath).href);
    let httpReaderCancelled = false;
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async () => new Response(
        new ReadableStream({
            start(controller) {
                controller.enqueue(new Uint8Array([1]));
            },
            cancel() {
                httpReaderCancelled = true;
            },
        }),
        {
            status: 200,
            headers: { "content-length": String(MAX_HTTP_RESPONSE_BYTES + 1) },
        },
    );
    try {
        const client = new HttpClient();
        const request = new HttpRequest("GET", "https://api.example.test", {});
        await assert.rejects(
            client.send(request),
            new RegExp(`exceeds the ${MAX_HTTP_RESPONSE_BYTES}-byte limit`),
        );
        assert.equal(httpReaderCancelled, true, "HTTP response reader must be cancelled");

        let fragmentedHttpCancelled = false;
        let httpChunksProduced = 0;
        globalThis.fetch = async () => new Response(
            new ReadableStream({
                pull(controller) {
                    httpChunksProduced += 1;
                    controller.enqueue(new Uint8Array([0x20]));
                },
                cancel() {
                    fragmentedHttpCancelled = true;
                },
            }),
            { status: 200 },
        );
        await assert.rejects(
            client.send(new HttpRequest("GET", "https://api.example.test", {})),
            new RegExp(`exceeds the ${MAX_RESPONSE_CHUNKS}-chunk limit`),
        );
        assert.ok(
            httpChunksProduced >= MAX_RESPONSE_CHUNKS + 1
            && httpChunksProduced <= MAX_RESPONSE_CHUNKS + 2,
            `unexpected HTTP chunk prefetch count: ${httpChunksProduced}`,
        );
        assert.equal(fragmentedHttpCancelled, true, "over-fragmented HTTP reader must be cancelled");
    } finally {
        globalThis.fetch = originalFetch;
    }
} finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
}

console.log("CodeQL HTTP security regression checks passed.");
