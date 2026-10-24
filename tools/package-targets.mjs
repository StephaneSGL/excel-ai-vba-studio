import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

export const PACKAGE_TARGETS = Object.freeze(['win32-x64', 'linux-x64']);

export function requestedTargets(args = process.argv.slice(2)) {
  if (args.length === 0) return [...PACKAGE_TARGETS];
  if (args.length !== 2 || args[0] !== '--target' || !PACKAGE_TARGETS.includes(args[1])) {
    throw new Error(`Usage: [--target ${PACKAGE_TARGETS.join('|')}]`);
  }
  return [args[1]];
}

/** Generate the target list from one common allowlist; never alter the checkout. */
export function withTargetIgnore(root, target, callback) {
  if (!PACKAGE_TARGETS.includes(target)) throw new Error(`Unsupported package target: ${target}`);
  const directory = mkdtempSync(join(tmpdir(), 'excel-ai-vsix-ignore-'));
  try {
    const ignoreFile = join(directory, '.vscodeignore');
    let common = readFileSync(resolve(root, '.vscodeignore'), 'utf8');
    // VSCE handles all negations after exclusions, regardless of line order.
    if (target === 'linux-x64') {
      common = common.split(/\r?\n/).filter(line => !['!scripts/**', '!bin/**'].includes(line.trim())).join('\n');
    }
    const targetExclusions = target === 'linux-x64' ? '\nscripts/**\nbin/**\n' : '';
    writeFileSync(ignoreFile, common + targetExclusions, { flag: 'wx' });
    return callback(ignoreFile);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
}
