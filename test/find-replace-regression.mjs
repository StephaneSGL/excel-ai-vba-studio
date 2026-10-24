import assert from 'node:assert/strict';
import { withModule } from './helpers/load-ts.mjs';

await withModule('src/react/view/excel/excel_find.ts', ({ findAllInSheets, findNextMatch, replaceCellText, replaceAllInSheets }) => {
  const sheet = (name, texts) => ({ name, rows: { len: 1, 0: { cells: Object.fromEntries(texts.map((text, i) => [i, { text }])) } } });
  const sheets = [sheet('One', ['Bonjour bonjour', 'a.b [x]', 'locked']), sheet('Two', ['BONJOUR'])];
  assert.equal(findAllInSheets(sheets, 'bonjour', { allSheets: true }).length, 2);
  assert.equal(replaceCellText(sheets, { sheetIndex: 0, ri: 0, ci: 0 }, 'bonjour', '$&', {}), true);
  assert.equal(sheets[0].rows[0].cells[0].text, '$& $&');
  assert.equal(replaceAllInSheets(sheets, 'a.b [x]', 'literal', {}), 1);
  assert.equal(replaceAllInSheets(sheets, 'BONJOUR', 'OK', { allSheets: true, matchCase: true, wholeCell: true }), 1);
  assert.equal(sheets[1].rows[0].cells[0].text, 'OK');
  assert.equal(replaceAllInSheets(sheets, '', 'bad', { allSheets: true }), 0);
  assert.equal(replaceCellText(sheets, { sheetIndex: -1, ri: 0, ci: 0 }, 'x', 'y', {}), false);
  assert.equal(replaceCellText(sheets, { sheetIndex: 0, ri: 0, ci: 1 }, 'lit', 'bad', { wholeCell: true }), false);
  sheets[0].rows[0].cells[2].editable = false;
  assert.equal(replaceAllInSheets(sheets, 'locked', 'bad', {}), 0);
  const multi = [sheet('A', ['x', 'x']), sheet('B', ['x'])];
  assert.deepEqual(findNextMatch(multi, 'x', { sheetIndex: 1, ri: 0, ci: 0 }, { allSheets: true }, 1), { sheetIndex: 0, ri: 0, ci: 0 });
  assert.equal(replaceAllInSheets(multi, 'x', 'y', { allSheets: true }), 3);
});

await withModule('src/react/util/loadOfficeContent.ts', ({ parseOfficeOpenPayload }) => {
  for (const input of [null, [], { ext: 42 }, { readOnly: 'false' }, { buffer: [-1] }, { buffer: [256] }, { backupSheets: {} }]) {
    assert.throws(() => parseOfficeOpenPayload(input), /Invalid/);
  }
  assert.equal(parseOfficeOpenPayload({ ext: 'xlsx', readOnly: true, buffer: [0, 255] }).readOnly, true);
});
const previousWindow = globalThis.window;
try {
  const keyHandlers = [];
  globalThis.window = { addEventListener(event, listener) { if (event === 'keydown') keyHandlers.push(listener); } };
  await withModule('src/react/util/vscode.ts', () => {
    for (const ctrlKey of [true, false]) {
      let cancelled = false;
      const event = { code: 'KeyV', ctrlKey, metaKey: !ctrlKey, altKey: false,
        preventDefault() { cancelled = true; } };
      keyHandlers.forEach(listener => listener(event));
      assert.equal(cancelled, false, 'the global host bridge must not cancel browser paste');
    }
  });
} finally {
  if (previousWindow === undefined) delete globalThis.window;
  else globalThis.window = previousWindow;
}
console.log('Find/replace, paste shortcut and webview message regressions passed.');
