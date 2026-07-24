import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import process from 'node:process';

const version = process.argv[2];
if (!version || !/^\d+\.\d+\.\d+$/.test(version)) {
  console.error('Usage: npm run release:prepare -- <major.minor.patch>');
  process.exit(1);
}

const root = resolve(import.meta.dirname, '..');
const manifestPath = resolve(root, 'package.json');
const lockPath = resolve(root, 'package-lock.json');
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));

if (manifest.version === version) {
  console.error(`package.json is already at version ${version}.`);
  process.exit(1);
}

manifest.version = version;
writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');

if (existsSync(lockPath)) {
  const lock = JSON.parse(readFileSync(lockPath, 'utf8'));
  lock.version = version;
  if (lock.packages?.['']) {
    lock.packages[''].version = version;
  }
  writeFileSync(lockPath, `${JSON.stringify(lock, null, 2)}\n`, 'utf8');
}

console.log(`Prepared version ${version}.`);
console.log('Next: update CHANGELOG.md, run npm run validate, commit, then create tag v' + version + '.');
