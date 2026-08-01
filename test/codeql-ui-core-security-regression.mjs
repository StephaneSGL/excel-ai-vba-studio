import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const temp = mkdtempSync(join(tmpdir(), 'codeql-ui-core-'));
const bundle = join(temp, 'security-regression.cjs');
const previousWindow = globalThis.window;

try {
  await build({
    stdin: {
      contents: [
        "export { sanitizeFileIconUrl as sanitizeTreeIconUrl } from './src/react/view/components/FileTypeIcon.tsx';",
        "export { sanitizeFileIconUrl as sanitizeArchiveIconUrl } from './src/react/view/compress/components/FileTypeIcon.tsx';",
        "export { normalizeOoxmlRelationshipTargets } from './src/react/view/excel/ooxml_namespace.ts';",
        "export { escapeCssSingleQuotedString } from './src/react/view/excel/x-spreadsheet/canvas/draw.js';",
        "export { escapeCssDoubleQuotedString } from './src/react/view/excel/x-spreadsheet/component/editor.js';",
      ].join('\n'),
      resolveDir: root,
      sourcefile: 'codeql-ui-core-security-entry.ts',
      loader: 'ts',
    },
    bundle: true,
    format: 'cjs',
    logLevel: 'silent',
    outfile: bundle,
    platform: 'node',
    target: 'node22',
  });

  globalThis.window = { devicePixelRatio: 1 };
  const {
    escapeCssDoubleQuotedString,
    escapeCssSingleQuotedString,
    normalizeOoxmlRelationshipTargets,
    sanitizeArchiveIconUrl,
    sanitizeTreeIconUrl,
  } = await import(pathToFileURL(bundle).href);

  const trustedIcon =
    'https://file+.vscode-resource.vscode-cdn.net/c%3A/extensions/icons/xlsx.svg';
  for (const sanitize of [sanitizeTreeIconUrl, sanitizeArchiveIconUrl]) {
    assert.equal(sanitize(trustedIcon), trustedIcon);
    assert.equal(sanitize('javascript:alert(1)'), null);
    assert.equal(sanitize('data:image/svg+xml,<svg/>'), null);
    assert.equal(sanitize('https://example.test/icon.svg'), null);
    assert.equal(
      sanitize('https://file+.vscode-resource.vscode-cdn.net.evil.test/icon.svg'),
      null,
    );
    assert.equal(
      sanitize('https://user@file+.vscode-resource.vscode-cdn.net/icon.svg'),
      null,
    );
    assert.equal(
      sanitize('https://file+.vscode-resource.vscode-cdn.net/icon.svg?raw=1'),
      null,
    );
    assert.equal(
      sanitize('https://file+.vscode-resource.vscode-cdn.net/icon.png'),
      null,
    );
  }

  assert.equal(escapeCssSingleQuotedString('\\'), '\\\\');
  assert.equal(escapeCssSingleQuotedString("'"), "\\'");
  assert.equal(escapeCssSingleQuotedString('A\nB'), 'A B');
  assert.equal(escapeCssDoubleQuotedString('\\'), '\\\\');
  assert.equal(escapeCssDoubleQuotedString('"'), '\\"');
  assert.equal(escapeCssDoubleQuotedString('A\rB'), 'A B');

  const namespace =
    'http://schemas.openxmlformats.org/package/2006/relationships';
  const validRelationships = [
    `<Relationships xmlns="${namespace}">`,
    '<Relationship Target="/xl/tables/table1.xml" Id="rId1"/>',
    '</Relationships>',
  ].join('');
  assert.equal(
    normalizeOoxmlRelationshipTargets(
      validRelationships,
      'xl/worksheets/_rels/sheet1.xml.rels',
    ).xml,
    validRelationships.replace('/xl/tables/table1.xml', '../tables/table1.xml'),
  );

  const prefixedRelationships = [
    `<pr:Relationships xmlns:pr="${namespace}">`,
    '<pr:Relationship Target="/xl/styles.xml" Id="rId1"/>',
    '</pr:Relationships>',
  ].join('');
  assert.equal(
    normalizeOoxmlRelationshipTargets(
      prefixedRelationships,
      'xl/worksheets/_rels/sheet1.xml.rels',
    ).xml,
    prefixedRelationships.replace('/xl/styles.xml', '../styles.xml'),
  );

  const substringNamespace = [
    `<Relationships xmlns="https://attacker.test/?next=${namespace}">`,
    '<Relationship Target="/xl/styles.xml" Id="rId1"/>',
    '</Relationships>',
  ].join('');
  assert.deepEqual(
    normalizeOoxmlRelationshipTargets(
      substringNamespace,
      'xl/worksheets/_rels/sheet1.xml.rels',
    ),
    { xml: substringNamespace, changed: false },
  );

  const wrongChildPrefix = [
    `<pr:Relationships xmlns:pr="${namespace}" xmlns:evil="urn:evil">`,
    '<evil:Relationship Target="/xl/styles.xml" Id="rId1"/>',
    '</pr:Relationships>',
  ].join('');
  assert.deepEqual(
    normalizeOoxmlRelationshipTargets(
      wrongChildPrefix,
      'xl/worksheets/_rels/sheet1.xml.rels',
    ),
    { xml: wrongChildPrefix, changed: false },
  );

  for (const relativePath of ['src/react/util/vscode.ts']) {
    const source = readFileSync(resolve(root, relativePath), 'utf8');
    assert.doesNotMatch(source, /events\s*\[\s*data\.type\s*\]/);
    assert.match(source, /new Map/);
    assert.match(source, /typeof eventHandler === ['"]function['"]/);
  }

  console.log('CodeQL UI/core security regressions passed.');
} finally {
  if (previousWindow === undefined) {
    delete globalThis.window;
  } else {
    globalThis.window = previousWindow;
  }
  rmSync(temp, { recursive: true, force: true });
}
