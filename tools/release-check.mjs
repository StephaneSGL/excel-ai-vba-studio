import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import process from 'node:process';

const manifest = JSON.parse(readFileSync(resolve(import.meta.dirname, '..', 'package.json'), 'utf8'));
const tag = process.argv[2] || process.env.GITHUB_REF_NAME;
const expectedTag = `v${manifest.version}`;

if (!tag) {
  console.error('Release tag is missing. Pass v<package-version> or set GITHUB_REF_NAME.');
  process.exit(1);
}

if (tag !== expectedTag) {
  console.error(`Release blocked: tag ${tag} does not match package.json version ${manifest.version} (expected ${expectedTag}).`);
  process.exit(1);
}

if (!/^\d+\.\d+\.\d+$/.test(manifest.version)) {
  console.error(`Release blocked: ${manifest.version} is not a Marketplace-compatible major.minor.patch version.`);
  process.exit(1);
}

console.log(`Release tag verified: ${tag}`);
