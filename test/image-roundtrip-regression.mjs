import assert from 'node:assert/strict';
import { withModule } from './helpers/load-ts.mjs';

await withModule({ source: "export * from './src/react/view/excel/excel_images.ts'; export { default as ExcelJS } from '@cweijan/exceljs';" }, async ({ ExcelJS, writeWorksheetImages, readWorksheetImages, readWorksheetBackgroundImage }) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Images');
  const base64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aGZkAAAAASUVORK5CYII=';
  writeWorksheetImages(sheet, workbook, [
    { id: 'one', imageId: 0, extension: 'png', base64, anchor: { col: 0.5, row: 1, width: 32, height: 24 } },
    { id: 'two', imageId: 1, extension: 'png', base64, anchor: { col: 3, row: 3, brCol: 5, brRow: 6 } },
  ], { imageId: 2, extension: 'png', base64 });
  const reread = new ExcelJS.Workbook();
  await reread.xlsx.load(await workbook.xlsx.writeBuffer());
  const loaded = reread.getWorksheet('Images');
  const images = readWorksheetImages(loaded, reread);
  assert.equal(images.length, 2);
  assert.equal(images[0].base64, base64);
  assert.equal(images[0].anchor.col, 0.5);
  assert.equal(images[0].anchor.width, 32);
  assert.equal(images[1].anchor.brRow, 6);
  assert.equal(readWorksheetBackgroundImage(loaded, reread).base64, base64);
});
console.log('Image/background OOXML round trips passed.');
