import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  mkdtemp,
  readFile,
  readdir,
  rm,
  writeFile,
} from 'node:fs/promises';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import JSZip from 'jszip';

const execFileAsync = promisify(execFile);
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const powershell = process.platform === 'win32'
  ? path.join(
      process.env.SystemRoot || 'C:\\Windows',
      'System32',
      'WindowsPowerShell',
      'v1.0',
      'powershell.exe',
    )
  : 'pwsh';

const sha256 = (value) => createHash('sha256').update(value).digest('hex');
const base64 = (value) => Buffer.from(value, 'utf8').toString('base64');

const contentTypesNamespace =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const relationshipsNamespace =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const officeDocumentRelationship =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const originRelationship =
  'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin';
const signatureRelationship =
  'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature';
const originContentType =
  'application/vnd.openxmlformats-package.digital-signature-origin';
const signatureContentType =
  'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml';

function encodeXml(value, encoding) {
  if (encoding === 'utf8') return value;
  const declared = value.replace(
    '<?xml version="1.0"?>',
    '<?xml version="1.0" encoding="UTF-16"?>',
  );
  const littleEndian = Buffer.from(declared, 'utf16le');
  if (encoding === 'utf16le') {
    return Buffer.concat([Buffer.from([0xff, 0xfe]), littleEndian]);
  }
  const bigEndian = Buffer.allocUnsafe(littleEndian.length);
  for (let index = 0; index < littleEndian.length; index += 2) {
    bigEndian[index] = littleEndian[index + 1];
    bigEndian[index + 1] = littleEndian[index];
  }
  return Buffer.concat([Buffer.from([0xfe, 0xff]), bigEndian]);
}

async function buildWorkbook({ signature = 'none', xmlEncoding = 'utf8' } = {}) {
  const zip = new JSZip();
  const alternative = signature === 'alternative';
  const ambiguous = signature === 'ambiguous';
  const emptyOrigin = signature === 'origin-empty';
  const emptyOriginRels = signature === 'origin-empty-rels';
  const malformedOriginRels = signature === 'origin-malformed-rels';
  const signed = signature === 'standard' || alternative;
  const hasOrigin =
    signed || ambiguous || emptyOrigin || emptyOriginRels || malformedOriginRels;
  const originPart = alternative ? '/security/origin.bin' : '/_xmlsignatures/origin.sigs';
  const signaturePart = alternative
    ? '/signatures/custom.xml'
    : '/_xmlsignatures/sig1.xml';
  zip.file(
    '[Content_Types].xml',
    encodeXml(`<?xml version="1.0"?><Types xmlns="${contentTypesNamespace}">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      (hasOrigin
        ? `<Override PartName="${originPart}" ContentType="${originContentType}"/>`
        : '') +
      (signed || ambiguous
        ? `<Override PartName="${signaturePart}" ContentType="${signatureContentType}"/>`
        : '') +
      '</Types>', xmlEncoding),
  );
  zip.file(
    '_rels/.rels',
    encodeXml(`<?xml version="1.0"?><Relationships xmlns="${relationshipsNamespace}">` +
      `<Relationship Id="rIdWorkbook" Type="${officeDocumentRelationship}" Target="xl/workbook.xml"/>` +
      (hasOrigin
        ? `<Relationship Id="rIdSignatureOrigin" Type="${originRelationship}" Target="${originPart.slice(1)}"/>`
        : '') +
      '</Relationships>', xmlEncoding),
  );
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>',
  );
  if (signed) {
    zip.file(originPart.slice(1), '');
    const originRels = alternative
      ? 'security/_rels/origin.bin.rels'
      : '_xmlsignatures/_rels/origin.sigs.rels';
    const target = alternative ? '../signatures/custom.xml' : 'sig1.xml';
    zip.file(
      originRels,
      encodeXml(`<?xml version="1.0"?><Relationships xmlns="${relationshipsNamespace}">` +
        `<Relationship Id="rIdSignature" Type="${signatureRelationship}" Target="${target}"/>` +
        '</Relationships>', xmlEncoding),
    );
    zip.file(
      signaturePart.slice(1),
      '<?xml version="1.0"?><Signature xmlns="http://www.w3.org/2000/09/xmldsig#"/>',
    );
  } else if (ambiguous) {
    // The relationship and content types claim a signature, but both targets are absent.
  } else if (emptyOrigin || emptyOriginRels || malformedOriginRels) {
    zip.file(originPart.slice(1), '');
    if (emptyOriginRels || malformedOriginRels) {
      zip.file(
        '_xmlsignatures/_rels/origin.sigs.rels',
        malformedOriginRels
          ? encodeXml('<?xml version="1.0"?><Relationships', xmlEncoding)
          : encodeXml(
              `<?xml version="1.0"?><Relationships xmlns="${relationshipsNamespace}"/>`,
              xmlEncoding,
            ),
      );
    }
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function corruptFirstLocalHeaderCrc(bytes) {
  const result = Buffer.from(bytes);
  const signature = Buffer.from([0x50, 0x4b, 0x03, 0x04]);
  const offset = result.indexOf(signature);
  assert.notEqual(offset, -1);
  result.writeUInt32LE((result.readUInt32LE(offset + 14) ^ 0xffffffff) >>> 0, offset + 14);
  return result;
}

function declareOversizedCompressedMetadata(bytes) {
  const result = Buffer.from(bytes);
  const signature = Buffer.from([0x50, 0x4b, 0x01, 0x02]);
  let offset = 0;
  while ((offset = result.indexOf(signature, offset)) >= 0) {
    const nameLength = result.readUInt16LE(offset + 28);
    const name = result.subarray(offset + 46, offset + 46 + nameLength).toString('utf8');
    if (name === '[Content_Types].xml') {
      result.writeUInt32LE(2 * 1024 * 1024 + 1, offset + 20);
      return result;
    }
    offset += 46 + nameLength;
  }
  assert.fail('Could not find [Content_Types].xml central-directory record.');
}

async function bundleModule(entryPoint, vscodeStub = undefined) {
  const result = await build({
    entryPoints: [entryPoint],
    bundle: true,
    format: 'cjs',
    platform: 'node',
    target: 'node20',
    write: false,
    logLevel: 'silent',
    plugins: vscodeStub
      ? [{
          name: 'vscode-stub',
          setup(buildApi) {
            buildApi.onResolve(
              { filter: /^vscode$/ },
              () => ({ path: 'vscode', namespace: 'stub' }),
            );
            buildApi.onLoad(
              { filter: /.*/, namespace: 'stub' },
              () => ({ contents: vscodeStub, loader: 'js' }),
            );
          },
        }]
      : [],
  });
  const module = { exports: {} };
  new Function(
    'module',
    'exports',
    'require',
    '__filename',
    '__dirname',
    result.outputFiles[0].text,
  )(
    module,
    module.exports,
    require,
    entryPoint,
    path.dirname(entryPoint),
  );
  return module.exports;
}

async function expectRefusal(executable, args, expectedText) {
  try {
    await execFileAsync(executable, args, {
      cwd: root,
      windowsHide: true,
      timeout: 30_000,
      maxBuffer: 1024 * 1024,
    });
  } catch (error) {
    const output = String(error.stdout ?? '') + '\n' + String(error.stderr ?? '');
    assert.match(output, expectedText);
    return output;
  }
  assert.fail(
    'Expected ' + path.basename(executable) + ' to refuse the signed package.',
  );
}

function powerShellUnsignedCommand(workbookPath) {
  const sharedScript = path.join(root, 'scripts', 'ooxml-package-signature.ps1');
  return (
    "Add-Type -AssemblyName System.IO.Compression.FileSystem; " +
    `. ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${base64(sharedScript)}'))); ` +
    `Assert-OoxmlPackageUnsigned ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${base64(workbookPath)}'))); ` +
    "[Console]::Out.Write('UNSIGNED')"
  );
}

async function expectPowerShellUnsigned(workbookPath) {
  const { stdout } = await execFileAsync(
    powershell,
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      powerShellUnsignedCommand(workbookPath),
    ],
    { cwd: root, windowsHide: true, timeout: 30_000, maxBuffer: 1024 * 1024 },
  );
  assert.equal(String(stdout), 'UNSIGNED');
}

async function snapshot(directory) {
  return (await readdir(directory)).sort();
}

const helperEntry = path.join(
  root,
  'src',
  'common',
  'ooxmlPackageSignature.ts',
);
const signatureHelper = await bundleModule(helperEntry);
assert.equal(typeof signatureHelper.hasOoxmlPackageSignature, 'function');
assert.equal(typeof signatureHelper.hasOoxmlPackageSignatureBytes, 'function');
assert.equal(
  typeof signatureHelper.assertOoxmlPackageUnsignedForMutation,
  'function',
);
assert.equal(
  typeof signatureHelper.assertOoxmlPackageUnsignedBytesForMutation,
  'function',
);

const tempRoot = await mkdtemp(
  path.join(os.tmpdir(), 'excel-ai-package-signature-lock-'),
);
try {
  const signedXlsx = path.join(tempRoot, 'signed-source.xlsx');
  const signedXlsm = path.join(tempRoot, 'signed-macro.xlsm');
  const signedXlsb = path.join(tempRoot, 'signed-binary.xlsb');
  const alternativeSignedXlsx = path.join(tempRoot, 'alt-source.xlsx');
  const alternativeSignedXlsm = path.join(tempRoot, 'alternative-signed.xlsm');
  const signedUtf16LeXlsx = path.join(tempRoot, 'signed-utf16le.xlsx');
  const signedUtf16BeXlsx = path.join(tempRoot, 'signed-utf16be.xlsx');
  const emptyOriginXlsm = path.join(tempRoot, 'empty-origin.xlsm');
  const emptyOriginRelsXlsm = path.join(tempRoot, 'empty-origin-rels.xlsm');
  const malformedOriginRelsXlsm = path.join(tempRoot, 'malformed-origin-rels.xlsm');
  const unsignedXlsx = path.join(tempRoot, 'unsigned.xlsx');
  const ambiguousXlsx = path.join(tempRoot, 'ambiguous.xlsx');
  const malformedXlsx = path.join(tempRoot, 'malformed.xlsx');
  const inconsistentZipXlsx = path.join(tempRoot, 'inconsistent-local-header.xlsx');
  const oversizedCompressedMetadataXlsx = path.join(
    tempRoot,
    'oversized-compressed-metadata.xlsx',
  );
  const signedBytes = await buildWorkbook({ signature: 'standard' });
  const alternativeSignedBytes = await buildWorkbook({ signature: 'alternative' });
  const signedUtf16LeBytes = await buildWorkbook({
    signature: 'alternative',
    xmlEncoding: 'utf16le',
  });
  const signedUtf16BeBytes = await buildWorkbook({
    signature: 'alternative',
    xmlEncoding: 'utf16be',
  });
  const emptyOriginBytes = await buildWorkbook({ signature: 'origin-empty' });
  const emptyOriginRelsBytes = await buildWorkbook({ signature: 'origin-empty-rels' });
  const malformedOriginRelsBytes = await buildWorkbook({
    signature: 'origin-malformed-rels',
  });
  const unsignedBytes = await buildWorkbook();
  const ambiguousBytes = await buildWorkbook({ signature: 'ambiguous' });
  await Promise.all([
    writeFile(signedXlsx, signedBytes, { flag: 'wx' }),
    writeFile(signedXlsm, signedBytes, { flag: 'wx' }),
    writeFile(signedXlsb, signedBytes, { flag: 'wx' }),
    writeFile(alternativeSignedXlsx, alternativeSignedBytes, { flag: 'wx' }),
    writeFile(alternativeSignedXlsm, alternativeSignedBytes, { flag: 'wx' }),
    writeFile(signedUtf16LeXlsx, signedUtf16LeBytes, { flag: 'wx' }),
    writeFile(signedUtf16BeXlsx, signedUtf16BeBytes, { flag: 'wx' }),
    writeFile(emptyOriginXlsm, emptyOriginBytes, { flag: 'wx' }),
    writeFile(emptyOriginRelsXlsm, emptyOriginRelsBytes, { flag: 'wx' }),
    writeFile(malformedOriginRelsXlsm, malformedOriginRelsBytes, { flag: 'wx' }),
    writeFile(unsignedXlsx, unsignedBytes, { flag: 'wx' }),
    writeFile(ambiguousXlsx, ambiguousBytes, { flag: 'wx' }),
    writeFile(malformedXlsx, Buffer.from('not-a-zip'), { flag: 'wx' }),
    writeFile(inconsistentZipXlsx, corruptFirstLocalHeaderCrc(unsignedBytes), { flag: 'wx' }),
    writeFile(
      oversizedCompressedMetadataXlsx,
      declareOversizedCompressedMetadata(unsignedBytes),
      { flag: 'wx' },
    ),
  ]);

  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(signedXlsx),
    true,
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(signedXlsb),
    true,
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(unsignedXlsx),
    false,
    'a clean OPC graph is unsigned',
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(alternativeSignedXlsm),
    true,
    'an arbitrary signature-part URI must be detected through OPC relationships',
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(signedUtf16LeXlsx),
    true,
    'UTF-16LE OPC metadata must be decoded before signature inspection',
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(signedUtf16BeXlsx),
    true,
    'UTF-16BE OPC metadata must be decoded before signature inspection',
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(emptyOriginXlsm),
    false,
    'a valid empty Digital Signature Origin without a relationships part is unsigned',
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignature(emptyOriginRelsXlsm),
    false,
    'a valid empty Digital Signature Origin with empty relationships is unsigned',
  );
  assert.equal(
    await signatureHelper.hasOoxmlPackageSignatureBytes(
      alternativeSignedBytes,
      'virtual.xlsm',
    ),
    true,
  );
  await assert.rejects(
    signatureHelper.assertOoxmlPackageUnsignedForMutation(signedXlsx),
    /signature numérique de package Office/,
  );
  await assert.rejects(
    signatureHelper.assertOoxmlPackageUnsignedForMutation(malformedXlsx),
    /n’a pas pu être vérifi/,
  );
  await assert.rejects(
    signatureHelper.assertOoxmlPackageUnsignedForMutation(ambiguousXlsx),
    /n’a pas pu être vérifi/,
  );
  await assert.rejects(
    signatureHelper.assertOoxmlPackageUnsignedForMutation(inconsistentZipXlsx),
    /n’a pas pu être vérifi/,
  );
  await assert.rejects(
    signatureHelper.assertOoxmlPackageUnsignedForMutation(malformedOriginRelsXlsm),
    /n’a pas pu être vérifi/,
    'a present but malformed origin relationships part remains fail-closed',
  );
  await assert.rejects(
    signatureHelper.assertOoxmlPackageUnsignedForMutation(
      oversizedCompressedMetadataXlsx,
    ),
    /n’a pas pu être vérifi/,
  );

  const virtualStore = new Map();
  globalThis.__excelAiVirtualSignatureStore = virtualStore;
  const officeContent = await bundleModule(
    path.join(root, 'src', 'provider', 'handlers', 'officeContent.ts'),
    `const store = globalThis.__excelAiVirtualSignatureStore;
     const key = (uri) => uri.toString();
     module.exports = {
       FilePermission: { Readonly: 1 },
       workspace: { fs: {
         stat: async (uri) => {
           if (!store.has(key(uri))) {
             const error = new Error('missing'); error.code = 'FileNotFound'; throw error;
           }
           const value = store.get(key(uri));
           return { size: value.byteLength, permissions: 0 };
         },
         readFile: async (uri) => {
           if (!store.has(key(uri))) {
             const error = new Error('missing'); error.code = 'FileNotFound'; throw error;
           }
           return store.get(key(uri));
         },
         writeFile: async (uri, bytes) => { store.set(key(uri), bytes); },
       } },
     };`,
  );
  assert.deepEqual(
    await officeContent.getEmbeddedSpreadsheetReadOnlyState({
      scheme: 'file',
      fsPath: signedXlsx,
    }),
    {
      readOnly: true,
      readOnlyReason: 'package-signature',
    },
  );

  const virtualSource = {
    scheme: 'mem',
    fsPath: '/virtual/source.xlsx',
    toString: () => 'mem:/virtual/source.xlsx',
  };
  const virtualTarget = {
    scheme: 'mem',
    fsPath: '/virtual/target.xlsx',
    toString: () => 'mem:/virtual/target.xlsx',
  };
  virtualStore.set(virtualSource.toString(), alternativeSignedBytes);
  virtualStore.set(virtualTarget.toString(), signedBytes);
  assert.deepEqual(
    await officeContent.getEmbeddedSpreadsheetReadOnlyState(virtualSource),
    { readOnly: true, readOnlyReason: 'package-signature' },
    'a writable virtual provider must still expose a signed package as read-only',
  );
  await assert.rejects(
    officeContent.assertUriOoxmlPackageUnsignedForMutation(virtualSource),
    /signature numérique de package Office/,
  );
  await assert.rejects(
    officeContent.assertExistingUriOoxmlPackageUnsignedForMutation(virtualTarget),
    /signature numérique de package Office/,
  );
  virtualStore.set(virtualSource.toString(), unsignedBytes);
  assert.deepEqual(
    await officeContent.getEmbeddedSpreadsheetReadOnlyState(virtualSource),
    { readOnly: false },
  );
  // Simulate a signature appearing after load: the pre-write recheck must catch it.
  virtualStore.set(virtualSource.toString(), alternativeSignedBytes);
  let virtualWrites = 0;
  try {
    await officeContent.assertUriOoxmlPackageUnsignedForMutation(virtualSource);
    virtualWrites += 1;
  } catch {
    // Expected fail-closed refusal before the provider write.
  }
  assert.equal(virtualWrites, 0);

  const runtimeState = {
    files: new Map(),
    reads: new Map(),
    writes: 0,
    saveTarget: undefined,
    readHook: undefined,
  };
  globalThis.__excelAiVirtualRuntime = runtimeState;
  const commonRuntime = await bundleModule(
    path.join(root, 'src', 'provider', 'compress', 'commonHandler.ts'),
    `const state = globalThis.__excelAiVirtualRuntime;
     const key = (uri) => uri.toString();
     const missing = () => { const error = new Error('missing'); error.code = 'FileNotFound'; return error; };
     class TabInputCustom {}
     const Uri = {
       file: (fsPath) => ({ scheme: 'file', fsPath, toString: () => 'file:' + fsPath }),
       joinPath: (uri, ...parts) => ({
         scheme: uri.scheme,
         fsPath: uri.fsPath.replace(/\\/[^/]*$/, '') + '/' + parts.at(-1),
         toString() { return this.scheme + ':' + this.fsPath; },
       }),
       parse: (value) => ({ toString: () => value }),
     };
     const fs = {
       stat: async (uri) => {
         const value = state.files.get(key(uri));
         if (!value) throw missing();
         return { size: value.byteLength, permissions: 0 };
       },
       readFile: async (uri) => {
         const uriKey = key(uri);
         const count = (state.reads.get(uriKey) || 0) + 1;
         state.reads.set(uriKey, count);
         if (state.readHook) return state.readHook(uri, count);
         const value = state.files.get(uriKey);
         if (!value) throw missing();
         return value;
       },
       writeFile: async (uri, bytes) => {
         state.writes += 1;
         state.files.set(key(uri), bytes);
       },
     };
     module.exports = {
       FilePermission: { Readonly: 1 },
       TabInputCustom,
       Uri,
       ViewColumn: { Active: 1, Beside: 2 },
       commands: { executeCommand: async () => undefined },
       env: { openExternal: async () => true },
       window: {
         showErrorMessage: async () => undefined,
         showWarningMessage: async () => undefined,
         showSaveDialog: async () => state.saveTarget,
         tabGroups: { activeTabGroup: { activeTab: undefined } },
       },
       workspace: { fs },
     };`,
  );
  const makeHandler = () => {
    const callbacks = new Map();
    return {
      callbacks,
      panel: { title: '', active: true, webview: {} },
      emit: () => undefined,
      on(name, callback) { callbacks.set(name, callback); return this; },
    };
  };

  const raceSource = {
    scheme: 'mem',
    fsPath: '/virtual/race-source.xlsx',
    toString: () => 'mem:/virtual/race-source.xlsx',
  };
  runtimeState.files.set(raceSource.toString(), unsignedBytes);
  runtimeState.readHook = (uri, count) =>
    uri.toString() === raceSource.toString() && count >= 2
      ? alternativeSignedBytes
      : runtimeState.files.get(uri.toString());
  const raceHandler = makeHandler();
  commonRuntime.handleCommonEvent(raceSource, raceHandler);
  await assert.rejects(
    raceHandler.callbacks.get('save')(Array.from(unsignedBytes)),
    /signature numérique de package Office/,
  );
  assert.equal(runtimeState.writes, 0, 'a signature appearing after load blocks the host write');

  const saveAsSource = {
    scheme: 'mem',
    fsPath: '/virtual/save-as-source.xlsx',
    toString: () => 'mem:/virtual/save-as-source.xlsx',
  };
  const saveAsTarget = {
    scheme: 'mem',
    fsPath: '/virtual/existing-signed-target.xlsx',
    toString: () => 'mem:/virtual/existing-signed-target.xlsx',
  };
  runtimeState.files.set(saveAsSource.toString(), unsignedBytes);
  runtimeState.files.set(saveAsTarget.toString(), alternativeSignedBytes);
  runtimeState.reads.clear();
  runtimeState.readHook = undefined;
  runtimeState.saveTarget = saveAsTarget;
  const saveAsHandler = makeHandler();
  commonRuntime.handleCommonEvent(saveAsSource, saveAsHandler);
  await assert.rejects(
    saveAsHandler.callbacks.get('saveAs')({
      content: Array.from(unsignedBytes),
      ext: 'xlsx',
    }),
    /signature numérique de package Office/,
  );
  assert.equal(runtimeState.writes, 0, 'a signed virtual Save As target is never overwritten');
  assert.deepEqual(
    await officeContent.getEmbeddedSpreadsheetReadOnlyState({
      scheme: 'file',
      fsPath: malformedXlsx,
    }),
    {
      readOnly: true,
      readOnlyReason: 'package-signature-verification',
    },
  );

  const [
    officeContentSource,
    commonHandler,
    excelView,
    nativeBridge,
    workbookService,
    writebackService,
    nativeScript,
    bootstrapScript,
    designerScript,
    sharedSignatureScript,
    nativeCli,
  ] = await Promise.all([
    readFile(path.join(root, 'src/provider/handlers/officeContent.ts'), 'utf8'),
    readFile(path.join(root, 'src/provider/compress/commonHandler.ts'), 'utf8'),
    readFile(path.join(root, 'src/react/view/excel/Excel.tsx'), 'utf8'),
    readFile(path.join(root, 'src/provider/nativeExcelBridge.ts'), 'utf8'),
    readFile(path.join(root, 'src/excelAiVbaStudio/workbookService.ts'), 'utf8'),
    readFile(path.join(root, 'src/excelAiVbaStudio/vbaWritebackService.ts'), 'utf8'),
    readFile(path.join(root, 'scripts/office-ai-apply-edits.ps1'), 'utf8'),
    readFile(path.join(root, 'scripts/prepare-macro-workbook.ps1'), 'utf8'),
    readFile(path.join(root, 'scripts/apply-vba-designer.ps1'), 'utf8'),
    readFile(path.join(root, 'scripts/ooxml-package-signature.ps1'), 'utf8'),
    readFile(path.join(root, 'native/vba-writeback/cli.py'), 'utf8'),
  ]);

  assert.ok(
    officeContentSource.indexOf('hasUriOoxmlPackageSignature(uri)') <
      officeContentSource.indexOf('isUriReadOnly(uri)'),
    'package signatures must take precedence over file-permission Save As',
  );
  assert.match(commonHandler, /assertUriOoxmlPackageUnsignedForMutation\(uri\)/);
  assert.match(
    commonHandler,
    /assertExistingUriOoxmlPackageUnsignedForMutation\(target\)/,
  );
  const sourceRecheckPositions = [
    ...commonHandler.matchAll(/assertUriOoxmlPackageUnsignedForMutation\(uri\)/g),
  ].map((match) => match.index);
  assert.ok(
    sourceRecheckPositions[0] < commonHandler.indexOf('workspace.fs.writeFile(uri, bytes)'),
    'source bytes must be rechecked before direct host writes',
  );
  assert.ok(
    sourceRecheckPositions[1] <
      commonHandler.indexOf('assertExistingUriOoxmlPackageUnsignedForMutation(target)') &&
    commonHandler.indexOf('assertExistingUriOoxmlPackageUnsignedForMutation(target)') <
      commonHandler.indexOf('workspace.fs.writeFile(target, bytes)'),
    'source and existing virtual Save As target must be checked before overwrite',
  );
  assert.match(commonHandler, /package-signature-verification/);
  assert.match(excelView, /blocksSaveAs[\s\S]+?package-signature/);
  assert.match(excelView, /packageSignatureWriteBlocked/);
  assert.match(
    nativeBridge,
    /assertOoxmlPackageUnsignedForMutation\(canonicalWorkbookPath\)/,
  );
  assert.match(
    workbookService,
    /writeVbaFromTool[\s\S]+?assertOoxmlPackageUnsignedForMutation/,
  );
  assert.match(
    workbookService,
    /designVbaFromTool[\s\S]+?assertOoxmlPackageUnsignedForMutation/,
  );
  assert.match(
    writebackService,
    /assertOoxmlPackageUnsignedForMutation\(workbookPath\)[\s\S]+?invokeHelper/,
  );
  for (const source of [nativeScript, bootstrapScript, designerScript]) {
    assert.match(source, /ooxml-package-signature[.]ps1/);
    assert.match(source, /Assert-OoxmlPackageUnsigned/);
  }
  assert.match(sharedSignatureScript, /digital-signature\/origin/);
  assert.match(sharedSignatureScript, /digital-signature\/signature/);
  assert.match(sharedSignatureScript, /OpcMaxEntries = 20000/);
  assert.doesNotMatch(sharedSignatureScript, /sig\[0-9\]/i);
  assert.match(nativeCli, /def assert_ooxml_package_unsigned/);
  assert.match(nativeCli, /ORIGIN_RELATIONSHIP_TYPE/);
  assert.match(nativeCli, /SIGNATURE_RELATIONSHIP_TYPE/);
  assert.match(nativeCli, /OOXML_PACKAGE_SIGNED/);

  await expectPowerShellUnsigned(emptyOriginXlsm);
  await expectPowerShellUnsigned(emptyOriginRelsXlsm);
  await expectRefusal(
    powershell,
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-Command',
      powerShellUnsignedCommand(malformedOriginRelsXlsm),
    ],
    /Package signature verification failed/,
  );

  const beforePowerShellTests = await snapshot(tempRoot);
  const nativePayloadPath = path.join(tempRoot, 'native-operations.json');
  await writeFile(nativePayloadPath, '{}', { flag: 'wx' });
  await expectRefusal(
    powershell,
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      path.join(root, 'scripts', 'office-ai-apply-edits.ps1'),
      '-WorkbookPath',
      alternativeSignedXlsm,
      '-OperationsPath',
      nativePayloadPath,
    ],
    /Office package signature detected/,
  );
  assert.deepEqual(
    await snapshot(tempRoot),
    [...beforePowerShellTests, 'native-operations.json'].sort(),
    'native XLSM refusal must not create an adjacent work file',
  );

  await expectRefusal(
    powershell,
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      path.join(root, 'scripts', 'prepare-macro-workbook.ps1'),
      '-WorkbookPathBase64',
      base64(alternativeSignedXlsx),
      '-ComponentFileBase64',
      base64('modSigned.bas'),
      '-SourceBase64',
      base64('Attribute VB_Name = "modSigned"\r\nPublic Sub Test()\r\nEnd Sub'),
    ],
    /Office package signature detected/,
  );
  assert.equal(
    (await snapshot(tempRoot)).some((name) => name.includes('bootstrap-')),
    false,
    'bootstrap refusal must happen before staging is created',
  );
  assert.equal(
    (await snapshot(tempRoot)).includes('alt-source.xlsm'),
    false,
    'bootstrap refusal must not create the target XLSM',
  );

  const helperExe = path.join(
    root,
    'bin',
    'win32-x64',
    'excel-ai-vba-writeback.exe',
  );
  const expectedWorkbookSha256 = sha256(alternativeSignedBytes);
  const designerRequestPath = path.join(tempRoot, 'designer-request.json');
  await writeFile(
    designerRequestPath,
    JSON.stringify({
      schemaVersion: 2,
      workbookPath: alternativeSignedXlsm,
      expectedWorkbookSha256,
      operations: [{ kind: 'createUserForm' }],
    }),
    { flag: 'wx' },
  );
  await expectRefusal(
    powershell,
    [
      '-NoLogo',
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy',
      'Bypass',
      '-File',
      path.join(root, 'scripts', 'apply-vba-designer.ps1'),
      '-RequestPathBase64',
      base64(designerRequestPath),
      '-HelperPathBase64',
      base64(helperExe),
    ],
    /Office package signature detected/,
  );
  assert.equal(
    (await snapshot(tempRoot)).some((name) => /^staging_.*[.]xlsm$/i.test(name)),
    false,
    'designer refusal must happen before staging is created',
  );

  if (process.platform === 'win32') {
  const writebackRequestPath = path.join(tempRoot, 'writeback-request.json');
  await writeFile(
    writebackRequestPath,
    JSON.stringify({
      schemaVersion: 1,
      workbookPath: alternativeSignedXlsm,
      expectedWorkbookSha256,
      patches: [{
        moduleName: 'modSigned',
        componentKind: 'module',
        source: 'Attribute VB_Name = "modSigned"\r\nPublic Sub Test()\r\nEnd Sub',
      }],
    }),
    { flag: 'wx' },
  );
  let writebackResult;
  try {
    await execFileAsync(helperExe, [writebackRequestPath], {
      cwd: tempRoot,
      windowsHide: true,
      timeout: 30_000,
      maxBuffer: 1024 * 1024,
    });
    assert.fail('Native write-back helper must reject a signed OOXML package.');
  } catch (error) {
    writebackResult = JSON.parse(String(error.stdout).trim());
  }
  assert.equal(writebackResult.ok, false);
  assert.equal(writebackResult.code, 'OOXML_PACKAGE_SIGNED');
  assert.match(writebackResult.message, /Office package signature detected/);
  assert.equal(
    (await snapshot(tempRoot)).some((name) =>
      name.startsWith('.alternative-signed.excel-ai-vba-')
    ),
    false,
    'native write-back refusal must happen before a work file is created',
  );
  assert.equal(
    (await readFile(alternativeSignedXlsm)).equals(alternativeSignedBytes),
    true,
    'every refusal must preserve the signed workbook byte-for-byte',
  );

  const malformedOriginRequestPath = path.join(
    tempRoot,
    'malformed-origin-request.json',
  );
  await writeFile(
    malformedOriginRequestPath,
    JSON.stringify({
      schemaVersion: 1,
      workbookPath: malformedOriginRelsXlsm,
      expectedWorkbookSha256: sha256(malformedOriginRelsBytes),
      patches: [{
        moduleName: 'modMalformedOrigin',
        componentKind: 'module',
        source:
          'Attribute VB_Name = "modMalformedOrigin"\r\nPublic Sub Test()\r\nEnd Sub',
      }],
    }),
    { flag: 'wx' },
  );
  let malformedOriginResult;
  try {
    await execFileAsync(helperExe, [malformedOriginRequestPath], {
      cwd: tempRoot,
      windowsHide: true,
      timeout: 30_000,
      maxBuffer: 1024 * 1024,
    });
    assert.fail('Malformed origin relationships must fail closed.');
  } catch (error) {
    malformedOriginResult = JSON.parse(String(error.stdout).trim());
  }
  assert.equal(malformedOriginResult.code, 'PACKAGE_SIGNATURE_UNVERIFIABLE');

  for (const [workbookPath, workbookBytes] of [
    [emptyOriginXlsm, emptyOriginBytes],
    [emptyOriginRelsXlsm, emptyOriginRelsBytes],
  ]) {
    const requestPath = `${workbookPath}.request.json`;
    await writeFile(
      requestPath,
      JSON.stringify({
        schemaVersion: 1,
        workbookPath,
        expectedWorkbookSha256: sha256(workbookBytes),
        patches: [{
          moduleName: 'modOriginOnly',
          componentKind: 'module',
          source:
            'Attribute VB_Name = "modOriginOnly"\r\nPublic Sub Test()\r\nEnd Sub',
        }],
      }),
      { flag: 'wx' },
    );
    let result;
    try {
      await execFileAsync(helperExe, [requestPath], {
        cwd: tempRoot,
        windowsHide: true,
        timeout: 30_000,
        maxBuffer: 1024 * 1024,
      });
      assert.fail('The synthetic workbook has no VBA project and cannot be edited.');
    } catch (error) {
      result = JSON.parse(String(error.stdout).trim());
    }
    assert.notEqual(result.code, 'OOXML_PACKAGE_SIGNED');
    assert.notEqual(result.code, 'PACKAGE_SIGNATURE_UNVERIFIABLE');
    assert.equal(
      (await readFile(workbookPath)).equals(workbookBytes),
      true,
      'an origin-only package remains byte-for-byte unchanged after downstream refusal',
    );
    assert.equal(
      (await snapshot(tempRoot)).some((name) =>
        name.startsWith(`.${path.parse(workbookPath).name}.excel-ai-vba-`)
      ),
      false,
      'origin-only analysis must not leak a write-back work file',
    );
  }
  }
} finally {
  await rm(tempRoot, { recursive: true, force: true });
}

console.log(
  'OOXML package-signature lock passed: grid, Save As, bootstrap, VBA write-back, UserForms, buttons, and ActiveX are fail-closed.',
);
