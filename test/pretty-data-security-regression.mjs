import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { build } from "esbuild";

const repositoryRoot = path.resolve(import.meta.dirname, "..");
const sourcePath = path.join(repositoryRoot, "src", "provider", "http", "utils", "pretty-data.ts");
const source = await fs.readFile(sourcePath, "utf8");

assert.match(source, /stripDelimitedComments/);
assert.doesNotMatch(source, /--\(\[\^\\-\]/, "the previous backtracking XML comment matcher must not return");
assert.doesNotMatch(source, /\(\[\^\*\]\|\[\\r\\n\]/, "the previous backtracking CSS comment matcher must not return");
assert.doesNotMatch(source, /ar\.indexOf\(line\)/, "XML formatting must use its linear loop index");

const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), "excel-ai-pretty-data-security-"));
const compiledModulePath = path.join(temporaryDirectory, "pretty-data.mjs");

try {
    await build({
        entryPoints: [sourcePath],
        outfile: compiledModulePath,
        bundle: true,
        format: "esm",
        platform: "node",
        target: "node20",
        logLevel: "silent",
    });

    const { PrettyData } = await import(pathToFileURL(compiledModulePath).href);
    const formatter = new PrettyData();

    assert.equal(
        formatter.xml("<root><item>one</item><item>two</item></root>"),
        ["<root>", "  <item>one</item>", "  <item>two</item>", "</root>"].join("\n"),
    );

    assert.equal(
        formatter.xmlmin("<root> <!-- first --> <item>value</item><!-- second --></root>"),
        "<root><item>value</item></root>",
    );
    assert.equal(
        formatter.xmlmin("<root><!-- retained", false),
        "<root><!-- retained",
        "unterminated input must be preserved",
    );
    assert.equal(
        formatter.xmlmin("<root><!-- retained --></root>", true),
        "<root><!-- retained --></root>",
    );

    assert.equal(formatter.cssmin("a { /* first */ color: red; /* second */ }"), "a {color: red;}");
    assert.equal(
        formatter.cssmin("a{/* retained", false),
        "a{/*retained",
        "unterminated input content must be preserved apart from minification",
    );
    assert.equal(formatter.cssmin("a{/* retained */ color:red}", true), "a{/*retained */color:red}");

    const adversarialXml = `<!--${"-".repeat(250_000)}`;
    const adversarialCss = `/*${"*".repeat(250_000)}`;
    assert.equal(formatter.xmlmin(adversarialXml), adversarialXml);
    assert.equal(formatter.cssmin(adversarialCss), adversarialCss);
} finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
}

console.log("Pretty-data security regression checks passed.");
