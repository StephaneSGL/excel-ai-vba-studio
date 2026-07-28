import assert from 'node:assert/strict';
import { mkdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';

const root = resolve(import.meta.dirname, '..');
const temporaryDirectory = join(
  tmpdir(),
  `excel-ai-vba-interactions-${process.pid}-${Date.now()}`,
);
await mkdir(temporaryDirectory, { recursive: false });

try {
  const output = join(temporaryDirectory, 'graph.mjs');
  await build({
    entryPoints: [
      resolve(root, 'src/excelAiVbaStudio/vbaInteractionGraph.ts'),
    ],
    outfile: output,
    bundle: true,
    format: 'esm',
    platform: 'node',
    logLevel: 'silent',
  });
  const { buildVbaInteractionGraph } = await import(
    `${pathToFileURL(output).href}?v=${Date.now()}`
  );

  const components = [
    {
      name: 'Module1',
      type: 'Standard module',
      source: [
        'Option Explicit',
        'Public Sub ShowOrders()',
        '    OrdersForm.Show',
        'End Sub',
        '',
        'Private Sub HiddenMacro()',
        '    OrdersForm.Show',
        'End Sub',
      ].join('\r\n'),
    },
    {
      name: 'Sheet1',
      type: 'Document module',
      source: [
        'Option Explicit',
        'Private Sub btnOpen_Click()',
        '    Call Module1.ShowOrders',
        'End Sub',
      ].join('\r\n'),
    },
    {
      name: 'Sheet2',
      type: 'Document module',
      source: [
        'Option Explicit',
        'Private Sub btnComplex_Click()',
        '    If Range("A1").Value Then',
        '        Call Module1.ShowOrders',
        '    End If',
        'End Sub',
        '',
        'Private Sub btnUnknown_Click()',
        '    MissingMacro',
        'End Sub',
      ].join('\r\n'),
    },
    {
      name: 'OrdersForm',
      type: 'UserForm',
      source: 'Option Explicit',
    },
  ];
  const graph = buildVbaInteractionGraph(components, {
    worksheetButtons: [
      {
        sheetName: 'Data',
        sheetCodeName: 'Sheet1',
        name: 'btnForm',
        caption: 'Open form',
        onAction: "'Book.xlsm'!Module1.ShowOrders",
        left: 10,
        top: 20,
        width: 100,
        height: 25,
      },
      {
        sheetName: 'Data',
        sheetCodeName: 'Sheet1',
        name: 'btnMissing',
        caption: 'Missing',
        onAction: 'MissingMacro',
        left: 10,
        top: 50,
        width: 100,
        height: 25,
      },
    ],
    worksheetActiveXControls: [
      {
        sheetName: 'Data',
        sheetCodeName: 'Sheet1',
        name: 'btnOpen',
        progId: 'Forms.CommandButton.1',
        caption: 'Open ActiveX',
        visible: true,
        left: 120,
        top: 20,
        width: 100,
        height: 25,
      },
      {
        sheetName: 'Other',
        sheetCodeName: 'Sheet2',
        name: 'btnComplex',
        progId: 'Forms.CommandButton.1',
        caption: 'Complex ActiveX',
        visible: true,
        left: 120,
        top: 50,
        width: 100,
        height: 25,
      },
      {
        sheetName: 'Other',
        sheetCodeName: 'Sheet2',
        name: 'btnUnknown',
        progId: 'Forms.CommandButton.1',
        caption: 'Missing ActiveX',
        visible: true,
        left: 120,
        top: 80,
        width: 100,
        height: 25,
      },
      {
        sheetName: 'Other',
        sheetCodeName: 'Sheet2',
        name: 'btnUnassigned',
        progId: 'Forms.CommandButton.1',
        caption: 'Unassigned ActiveX',
        visible: true,
        left: 120,
        top: 110,
        width: 100,
        height: 25,
      },
    ],
  });

  assert.deepEqual(
    graph.macros.map(({ qualifiedName }) => qualifiedName),
    ['Module1.ShowOrders'],
  );
  assert.deepEqual(graph.macros[0].userFormsOpened, ['OrdersForm']);
  assert.deepEqual(graph.userForms, ['OrdersForm']);
  assert.equal(graph.worksheetButtons.length, 2);
  assert.equal(graph.worksheetActiveXControls.length, 4);
  assert.equal(graph.relationships[0].resolution, 'resolved');
  assert.equal(graph.relationships[0].macroName, 'Module1.ShowOrders');
  assert.deepEqual(graph.relationships[0].userFormsOpened, ['OrdersForm']);
  assert.equal(graph.relationships[1].resolution, 'missing-macro');
  assert.equal(graph.relationships[2].resolution, 'resolved');
  assert.equal(graph.relationships[2].macroName, 'Module1.ShowOrders');
  assert.equal(
    graph.worksheetActiveXControls[0].macroName,
    'Module1.ShowOrders',
  );
  assert.equal(graph.relationships[3].resolution, 'complex-event');
  assert.equal(graph.relationships[4].resolution, 'missing-macro');
  assert.equal(graph.relationships[5].resolution, 'unassigned');

  console.log(
    'VBA interaction graph passed: public macros, Form buttons, ActiveX Click, and UserForm.Show resolved without execution.',
  );
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true });
}
