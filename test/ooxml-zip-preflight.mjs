import assert from 'node:assert/strict';
import { build } from 'esbuild';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import JSZip from 'jszip';

const root = resolve(import.meta.dirname, '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'excel-ooxml-preflight-'));
const outputPath = join(buildDirectory, 'ooxml-zip-preflight.mjs');

async function fixture(entries) {
  const zip = new JSZip();
  for (const [name, content] of entries) zip.file(name, content);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

try {
  await build({
    entryPoints: [join(root, 'src', 'react', 'view', 'excel', 'ooxml-zip-preflight.ts')],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'browser',
    logLevel: 'silent',
    tsconfig: join(root, 'tsconfig.json'),
  });
  const {
    inspectOoxmlZipCentralDirectory,
    validateOoxmlZipInflationBounds,
    DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS,
  } =
    await import(`${pathToFileURL(outputPath).href}?cache=${Date.now()}`);

  const safe = await fixture([
    ['[Content_Types].xml', '<Types/>'],
    ['xl/workbook.xml', '<workbook/>'],
    ['xl/media/image.png', new Uint8Array([1, 2, 3])],
  ]);
  const inspected = inspectOoxmlZipCentralDirectory(safe);
  assert.equal([...inspected.entries.values()].filter(entry => !entry.directory).length, 3);
  await validateOoxmlZipInflationBounds(safe, inspected);
  assert.ok(inspected.totalXmlBytes > 0);
  assert.ok(inspected.totalUncompressedBytes >= inspected.totalXmlBytes + 3);

  assert.throws(
    () => inspectOoxmlZipCentralDirectory(safe, {
      ...DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS,
      maxEntries: 2,
    }),
    /plus de 2 entrées ZIP/,
  );

  const forged = await fixture([['xl/workbook.xml', 'A'.repeat(4096)]]);
  const forgedView = new DataView(forged.buffer, forged.byteOffset, forged.byteLength);
  const encodedName = new TextEncoder().encode('xl/workbook.xml');
  const matchesNameAt = offset => encodedName.every((byte, index) => forged[offset + index] === byte);
  let patchedLocal = false;
  let patchedCentral = false;
  for (let offset = 0; offset + 46 < forged.byteLength; offset += 1) {
    const signature = forgedView.getUint32(offset, true);
    if (signature === 0x04034b50) {
      const nameLength = forgedView.getUint16(offset + 26, true);
      if (nameLength === encodedName.length && matchesNameAt(offset + 30)) {
        forgedView.setUint32(offset + 22, 8, true);
        patchedLocal = true;
      }
    }
    if (signature === 0x02014b50) {
      const nameLength = forgedView.getUint16(offset + 28, true);
      if (nameLength === encodedName.length && matchesNameAt(offset + 46)) {
        forgedView.setUint32(offset + 24, 8, true);
        patchedCentral = true;
      }
    }
  }
  assert.ok(patchedLocal && patchedCentral, 'forged fixture metadata was not patched');
  const forgedInspected = inspectOoxmlZipCentralDirectory(forged);
  await assert.rejects(
    validateOoxmlZipInflationBounds(forged, forgedInspected),
    /taille décompressée réelle supérieure/,
  );
  assert.throws(
    () => inspectOoxmlZipCentralDirectory(safe, {
      ...DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS,
      maxXmlEntryBytes: 2,
    }),
    /partie XML trop volumineuse/,
  );
  assert.throws(
    () => inspectOoxmlZipCentralDirectory(safe, {
      ...DEFAULT_OOXML_ZIP_PREFLIGHT_LIMITS,
      maxTotalXmlBytes: 8,
    }),
    /taille XML décompressée totale/,
  );

  const collision = await fixture([
    ['xl/workbook.xml', '<workbook/>'],
    ['XL/WORKBOOK.XML', '<workbook/>'],
  ]);
  assert.throws(
    () => inspectOoxmlZipCentralDirectory(collision),
    /collision de casse/,
  );

  console.log('OOXML ZIP preflight passed: metadata and real streamed inflation bounds enforced before JSZip.');
} finally {
  await rm(buildDirectory, { recursive: true, force: true });
}
