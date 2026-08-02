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
const manifestSource = readFileSync(manifestPath, 'utf8');
const manifest = JSON.parse(manifestSource);

function replaceFirstVersion(source, currentVersion, nextVersion, label) {
  const escapedVersion = currentVersion.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(`(^\\s*"version"\\s*:\\s*")${escapedVersion}("\\s*[,}])`, 'm');
  if (!pattern.test(source)) {
    throw new Error(`Could not locate ${label} version ${currentVersion}.`);
  }
  return source.replace(pattern, `$1${nextVersion}$2`);
}

if (manifest.version === version) {
  console.error(`package.json is already at version ${version}.`);
  process.exit(1);
}

const updatedManifestSource = replaceFirstVersion(
  manifestSource,
  manifest.version,
  version,
  'package.json',
);
writeFileSync(manifestPath, updatedManifestSource, 'utf8');

if (existsSync(lockPath)) {
  const lockSource = readFileSync(lockPath, 'utf8');
  const lock = JSON.parse(lockSource);
  if (lock.version !== manifest.version || lock.packages?.['']?.version !== manifest.version) {
    throw new Error('package-lock.json root versions do not match package.json.');
  }

  let updatedLockSource = replaceFirstVersion(
    lockSource,
    manifest.version,
    version,
    'package-lock.json top-level',
  );
  updatedLockSource = replaceFirstVersion(
    updatedLockSource,
    manifest.version,
    version,
    'package-lock.json root package',
  );
  writeFileSync(lockPath, updatedLockSource, 'utf8');
}

console.log(`Prepared version ${version}.`);
console.log('Next: update CHANGELOG.md, run npm run validate, commit, then create tag v' + version + '.');
