import { readFileSync, readdirSync, statSync } from 'node:fs';
import { resolve } from 'node:path';

const root = resolve(import.meta.dirname, '..');
const excelSource = readFileSync(
  resolve(root, 'src', 'react', 'view', 'excel', 'Excel.tsx'),
  'utf8',
);
const readerSource = readFileSync(
  resolve(root, 'src', 'react', 'view', 'excel', 'excel_reader.ts'),
  'utf8',
);

const requiredPatterns = [
  [excelSource, /import\(['"]\.\/excel_writer\.ts['"]\)/, 'Excel writer must be lazy-loaded'],
  [excelSource, /import\(['"]\.\/excel_reader\.ts['"]\)/, 'Excel reader must be lazy-loaded'],
  [readerSource, /import\(['"]@cweijan\/exceljs['"]\)/, 'ExcelJS must be lazy-loaded'],
  [readerSource, /import\(['"]jszip['"]\)/, 'JSZip must be lazy-loaded'],
  [readerSource, /import\(['"]xlsx['"]\)/, 'SheetJS must be lazy-loaded'],
  [readerSource, /import\(['"]udsv['"]\)/, 'uDSV must be lazy-loaded'],
];

for (const [source, pattern, message] of requiredPatterns) {
  if (!pattern.test(source)) {
    throw new Error(message);
  }
}

if (/^import\s+.*from\s+['"]\.\/excel_writer\.ts['"];?$/m.test(excelSource)) {
  throw new Error('Excel writer was moved back into the opening path');
}

if (/^import\s+.*from\s+['"]\.\/excel_reader\.ts['"];?$/m.test(excelSource)) {
  throw new Error('Excel reader was moved back into the opening path');
}

const assetsDir = resolve(root, 'out', 'webview', 'assets');
const entryFiles = readdirSync(assetsDir)
  .filter(name => /^index-[A-Za-z0-9_-]+\.js$/.test(name));

if (entryFiles.length !== 1) {
  throw new Error(`Expected one webview entry bundle, found ${entryFiles.length}`);
}

const entryBytes = statSync(resolve(assetsDir, entryFiles[0])).size;
const maxEntryBytes = 350 * 1024;
if (entryBytes > maxEntryBytes) {
  throw new Error(
    `Webview entry bundle is ${(entryBytes / 1024).toFixed(1)} KiB; limit is ${maxEntryBytes / 1024} KiB`,
  );
}

console.log(
  `Performance validation passed: initial webview JavaScript ${(entryBytes / 1024).toFixed(1)} KiB; heavy spreadsheet engines are lazy-loaded.`,
);
