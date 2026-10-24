import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const buildDirectory = await mkdtemp(join(tmpdir(), 'clipboard-html-security-'));
const outputPath = join(buildDirectory, 'clipboard-html.mjs');

function textNode(value) {
  return {
    nodeType: 3,
    nodeValue: value,
    childNodes: [],
  };
}

function element(tagName, children = []) {
  return {
    nodeType: 1,
    tagName,
    childNodes: children,
  };
}

function cell(value, attributes = {}, children = undefined) {
  return {
    ...element('TD', children ?? [textNode(value)]),
    getAttribute(name) {
      return Object.hasOwn(attributes, name) ? String(attributes[name]) : null;
    },
  };
}

function table(rowCells) {
  const result = element('TABLE');
  result.rows = rowCells.map((cells) => {
    const row = element('TR', cells);
    row.cells = cells;
    row.closest = selector => (selector === 'table' ? result : null);
    return row;
  });
  result.childNodes = result.rows;
  return result;
}

const computedStyle = Object.freeze({
  fontFamily: 'Arial',
  fontSize: '11pt',
  fontWeight: '400',
  fontStyle: 'normal',
  color: 'rgb(0, 0, 0)',
  backgroundColor: 'rgba(0, 0, 0, 0)',
  textDecorationLine: 'none',
  textAlign: 'left',
  verticalAlign: 'middle',
  whiteSpace: 'normal',
  borderTop: 'none',
  borderRight: 'none',
  borderBottom: 'none',
  borderLeft: 'none',
});

const previousWindow = globalThis.window;
try {
  await build({
    entryPoints: [join(
      root,
      'src',
      'react',
      'view',
      'excel',
      'x-spreadsheet',
      'core',
      'clipboard_html.js',
    )],
    outfile: outputPath,
    bundle: true,
    format: 'esm',
    platform: 'browser',
    logLevel: 'silent',
  });
  globalThis.window = { getComputedStyle: () => computedStyle };
  const {
    HtmlClipboardLimitError,
    isHtmlClipboardDestinationInBounds,
    parseHtmlClipboardTable,
  } = await import(`${pathToFileURL(outputPath).href}?cache=${Date.now()}`);

  const limits = {
    maxRows: 10,
    maxColumns: 10,
    maxCells: 12,
    maxDomNodes: 100,
  };

  const rectangular = parseHtmlClipboardTable(table([
    [cell('A'), cell('B')],
    [cell('C'), cell('D')],
  ]), limits);
  assert.deepEqual(rectangular.rows.map(row => row.map(item => item.text)), [
    ['A', 'B'],
    ['C', 'D'],
  ]);
  assert.equal(rectangular.height, 2);
  assert.equal(rectangular.width, 2);
  assert.deepEqual(rectangular.merges, []);

  const merged = parseHtmlClipboardTable(table([
    [cell('merged', { rowspan: 2, colspan: 3 }), cell('right')],
    [cell('below-right')],
  ]), { ...limits, maxCells: 8 });
  assert.deepEqual(merged.rows[0][0].merge, [1, 2]);
  assert.deepEqual(merged.merges, [{ sri: 0, sci: 0, eri: 1, eci: 2 }]);
  assert.equal(merged.rows[0][3].text, 'right');
  assert.equal(merged.rows[1][3].text, 'below-right');
  assert.equal(merged.height, 2);
  assert.equal(merged.width, 4);

  const exactBoundary = parseHtmlClipboardTable(table([
    [cell('exact', { rowspan: 2, colspan: 3 })],
  ]), { ...limits, maxCells: 6 });
  assert.equal(exactBoundary.height * exactBoundary.width, 6);

  let styleReads = 0;
  globalThis.window = {
    getComputedStyle: () => {
      styleReads += 1;
      return computedStyle;
    },
  };
  assert.throws(
    () => parseHtmlClipboardTable(table([
      [cell('blocked', { rowspan: 3, colspan: 3 })],
    ]), { ...limits, maxCells: 8 }),
    HtmlClipboardLimitError,
  );
  assert.equal(styleReads, 0, 'span overflow must be rejected before style work');

  assert.throws(
    () => parseHtmlClipboardTable(table([
      [cell('one', { rowspan: 2, colspan: 2 }), cell('two', { rowspan: 2, colspan: 2 })],
    ]), { ...limits, maxCells: 7 }),
    /expanded-cell budget/,
  );

  assert.throws(
    () => parseHtmlClipboardTable(table([
      [cell('wide', { colspan: 4 })],
      [cell('narrow')],
      [cell('narrow')],
    ]), { ...limits, maxCells: 8 }),
    /bounding-cell budget/,
  );

  for (const unsafe of ['999999999999999999999999', '11']) {
    assert.throws(
      () => parseHtmlClipboardTable(table([[cell('unsafe', { colspan: unsafe })]]), limits),
      HtmlClipboardLimitError,
    );
  }
  const normalizedInvalid = parseHtmlClipboardTable(table([[
    cell('invalid defaults to one', { colspan: '-4', rowspan: 'garbage' }),
  ]]), limits);
  assert.equal(normalizedInvalid.width, 1);
  assert.equal(normalizedInvalid.height, 1);

  const nestedCell = cell('outer', {}, [
    textNode('outer '),
    table([[cell('nested')]]),
  ]);
  const nested = parseHtmlClipboardTable(table([[nestedCell]]), limits);
  assert.equal(nested.rows.length, 1);
  assert.equal(nested.width, 1);
  assert.equal(nested.rows[0][0].text, 'outer nested');

  let deep = textNode('leaf');
  for (let index = 0; index < 20; index += 1) deep = element('SPAN', [deep]);
  assert.throws(
    () => parseHtmlClipboardTable(table([[cell('', {}, [deep])]]), {
      ...limits,
      maxDomNodes: 10,
    }),
    /DOM traversal budget/,
  );

  assert.equal(
    isHtmlClipboardDestinationInBounds(1_048_575, 16_383, 1, 1),
    true,
  );
  assert.equal(
    isHtmlClipboardDestinationInBounds(1_048_575, 16_383, 2, 1),
    false,
  );
  assert.equal(
    isHtmlClipboardDestinationInBounds(1_048_575, 16_383, 1, 2),
    false,
  );

  console.log('HTML clipboard security regression passed: spans, aggregate work, bounds, nesting, and DOM traversal are bounded.');
} finally {
  if (previousWindow === undefined) delete globalThis.window;
  else globalThis.window = previousWindow;
  await rm(buildDirectory, { recursive: true, force: true });
}
