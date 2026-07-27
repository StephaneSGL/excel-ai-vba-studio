import assert from 'node:assert/strict';
import {
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import JSZip from 'jszip';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const fixture = resolve(
  root,
  'test/fixtures/Excel-AI-VBA-Studio-Demo-base.xlsx',
);
const temp = mkdtempSync(join(tmpdir(), 'xlsx-prefix-regression-'));
const bundle = join(temp, 'excel-reader.cjs');

assert.ok(existsSync(fixture), `Missing fixture: ${fixture}`);
assert.equal(statSync(fixture).size, 19741);

const fixtureBuffer = readFileSync(fixture);
const fixtureZip = await JSZip.loadAsync(fixtureBuffer);
const workbookEntry = fixtureZip.file('xl/workbook.xml');
assert.ok(workbookEntry, 'Fixture has no xl/workbook.xml');
assert.match(await workbookEntry.async('string'), /<x:workbook\b/);

try {
  await build({
    stdin: {
      contents: [
        "export { readExcel } from './src/react/view/excel/excel_reader.ts';",
        "export { normalizeOoxmlRelationshipTargets, normalizeSpreadsheetMlElementPrefixes } from './src/react/view/excel/ooxml_namespace.ts';",
      ].join('\n'),
      resolveDir: root,
      sourcefile: 'xlsx-prefix-regression-entry.ts',
      loader: 'ts',
    },
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile: bundle,
    platform: 'node',
    target: 'node22',
  });
  const {
    normalizeOoxmlRelationshipTargets,
    normalizeSpreadsheetMlElementPrefixes,
    readExcel,
  } = await import(pathToFileURL(bundle).href);

  const plain =
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>';
  assert.deepEqual(normalizeSpreadsheetMlElementPrefixes(plain), {
    xml: plain,
    changed: false,
  });

  const prefixed = [
    '<?probe <x:fake>?>',
    '<!-- <x:fake> -->',
    '<![CDATA[<x:fake>]]>',
    '<x:workbook',
    ' xmlns:x = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"',
    ' r:id="rId1">',
    '<x:sheet xdr:col="3"><xdr:col>4</xdr:col></x:sheet>',
    '</x:workbook>',
  ].join('');
  const normalized = normalizeSpreadsheetMlElementPrefixes(prefixed);
  assert.equal(normalized.changed, true);
  assert.match(normalized.xml, /^<\?probe <x:fake>\?>/);
  assert.match(normalized.xml, /<!-- <x:fake> -->/);
  assert.match(normalized.xml, /<!\[CDATA\[<x:fake>]]>/);
  assert.match(normalized.xml, /<workbook\b/);
  assert.match(normalized.xml, /<sheet xdr:col="3">/);
  assert.match(normalized.xml, /<xdr:col>4<\/xdr:col>/);
  assert.match(normalized.xml, /r:id="rId1"/);
  assert.doesNotMatch(normalized.xml, /<\/?x:(?:workbook|sheet)\b/);

  const otherNamespace = '<x:root xmlns:x="urn:other"><x:child/></x:root>';
  assert.deepEqual(normalizeSpreadsheetMlElementPrefixes(otherNamespace), {
    xml: otherNamespace,
    changed: false,
  });

  const relationships = [
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Target="/xl/tables/table1.xml" Id="rId1"/>',
    '<Relationship Target="/xl/media/image1.png" TargetMode="External" Id="rId2"/>',
    '</Relationships>',
  ].join('');
  const normalizedRelationships = normalizeOoxmlRelationshipTargets(
    relationships,
    'xl/worksheets/_rels/sheet1.xml.rels',
  );
  assert.equal(normalizedRelationships.changed, true);
  assert.match(normalizedRelationships.xml, /Target="\.\.\/tables\/table1\.xml"/);
  assert.match(
    normalizedRelationships.xml,
    /Target="\/xl\/media\/image1\.png" TargetMode="External"/,
  );

  const arrayBuffer = fixtureBuffer.buffer.slice(
    fixtureBuffer.byteOffset,
    fixtureBuffer.byteOffset + fixtureBuffer.byteLength,
  );
  const workbook = await readExcel(arrayBuffer);
  assert.deepEqual(
    workbook.sheets.map((sheet) => sheet.name),
    ['Dashboard', 'Orders', 'Inventory', 'Customers', 'Settings'],
  );
  assert.equal(workbook.sheets.length, 5);

  console.log(
    'XLSX namespace-prefix regression passed: exact demo fixture opens with 5 sheets.',
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
