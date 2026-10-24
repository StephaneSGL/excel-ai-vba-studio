import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import JSZip from 'jszip';

const require = createRequire(import.meta.url);
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const helperPath = join(root, 'src', 'common', 'ooxmlPackageSignature.ts');
const contentTypesNamespace =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const relationshipsNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';

async function loadHelper() {
  const result = await build({
    entryPoints: [helperPath],
    bundle: true,
    format: 'cjs',
    platform: 'node',
    target: 'node20',
    write: false,
    logLevel: 'silent',
  });
  const module = { exports: {} };
  new Function('module', 'exports', 'require', result.outputFiles[0].text)(
    module,
    module.exports,
    require,
  );
  return module.exports;
}

async function packageBytes(contentTypes, relationships) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', relationships);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const helperSource = await readFile(helperPath, 'utf8');
assert.doesNotMatch(
  helperSource,
  /new Map\(stack[.]at\(-1\)[?][.]namespaces/,
  'OPC metadata parsing must not clone every inherited namespace map',
);
assert.match(helperSource, /restoreNamespaces\(current[.]namespaceChanges\)/);
assert.match(helperSource, /previousValue === undefined/);
assert.doesNotMatch(
  helperSource,
  /attributeText = attributeText[.]slice/,
  'attribute parsing must not repeatedly copy the unparsed suffix',
);

const helper = await loadHelper();

const scopedContentTypes = [
  `<ct:Types xmlns:ct="${contentTypesNamespace}" xmlns:s="${contentTypesNamespace}">`,
  '<ct:Default xmlns:s="urn:temporary-shadow" Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
  '<s:Default Extension="xml" ContentType="application/xml"></s:Default>',
  '</ct:Types>',
].join('');
const scopedRelationships = [
  `<r:Relationships xmlns:r="${relationshipsNamespace}" xmlns:s="${relationshipsNamespace}">`,
  '<r:Relationship xmlns:s="urn:temporary-shadow" Id="rId1" Type="urn:safe:first" Target="one.xml"/>',
  '<s:Relationship Id="rId2" Type="urn:safe:second" Target="two.xml"></s:Relationship>',
  '</r:Relationships>',
].join('');
const scopedBytes = await packageBytes(scopedContentTypes, scopedRelationships);
assert.equal(await helper.hasOoxmlPackageSignatureBytes(scopedBytes, 'scoped.xlsx'), false);
assert.equal(await helper.hasOoxmlXlmMacroSheetsBytes(scopedBytes, 'scoped.xlsx'), false);

const namespaceCount = 1_200;
const siblingCount = 1_200;
const namespaceDeclarations = Array.from(
  { length: namespaceCount },
  (_, index) => `xmlns:n${index}="urn:unused:${index}"`,
).join(' ');
const defaultDeclarations = Array.from(
  { length: siblingCount },
  (_, index) => `<ct:Default Extension="x${index}" ContentType="application/x-${index}"/>`,
).join('');
const relationshipDeclarations = Array.from(
  { length: siblingCount },
  (_, index) => `<r:Relationship Id="rId${index}" Type="urn:safe:${index}" Target="part-${index}.xml"/>`,
).join('');
const stressBytes = await packageBytes(
  `<ct:Types xmlns:ct="${contentTypesNamespace}" ${namespaceDeclarations}>${defaultDeclarations}</ct:Types>`,
  `<r:Relationships xmlns:r="${relationshipsNamespace}" ${namespaceDeclarations}>${relationshipDeclarations}</r:Relationships>`,
);
assert.equal(await helper.hasOoxmlPackageSignatureBytes(stressBytes, 'stress.xlsx'), false);
assert.equal(await helper.hasOoxmlXlmMacroSheetsBytes(stressBytes, 'stress.xlsx'), false);

const nestedBytes = await packageBytes(
  `<Types xmlns="${contentTypesNamespace}"><Default Extension="xml" ContentType="application/xml"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Default></Types>`,
  `<Relationships xmlns="${relationshipsNamespace}"/>`,
);
await assert.rejects(
  helper.assertOoxmlPackageUnsignedBytesForMutation(nestedBytes, 'nested.xlsx'),
  /n’a pas pu être vérifi/,
);

console.log(
  'OOXML namespace security regression passed: inherited scopes restore correctly and namespace-heavy siblings remain linear.',
);
