import fs from 'node:fs';
import crypto from 'node:crypto';
import process from 'node:process';
import XLSX from 'xlsx';

/**
 * Read all streams AND storages from a Compound File Binary (vbaProject.bin).
 * Returns full-path-indexed maps so callers can detect UserForms
 * from their project-root designer storages and form-specific child streams.
 */
function readCfb(filePath) {
  const compoundFile = XLSX.CFB.read(fs.readFileSync(filePath), {
    type: 'buffer',
  });
  const streams = new Map();
  const storages = [];

  for (let index = 0; index < compoundFile.FileIndex.length; index += 1) {
    const entry = compoundFile.FileIndex[index];
    if (!entry) continue;

    const path = compoundFile.FullPaths?.[index] ?? entry.name ?? `entry-${index}`;

    if (entry.type === 1) {
      storages.push(path);
    } else if (entry.type === 2) {
      const content = Buffer.from(entry.content ?? []);
      streams.set(path, {
        bytes: content.length,
        sha256: crypto.createHash('sha256').update(content).digest('hex'),
      });
    }
  }

  return { streams, storages: [...new Set(storages)].sort() };
}

/**
 * Build a conservative VBA inventory from CFB paths.
 *
 * Module kinds are stored in the compressed VBA/dir stream. This helper does
 * not parse that binary record stream, so it deliberately does not claim that
 * a module stream is standard, class, or document code. Designer storages are
 * detectable structurally because MS-OVBA stores them outside the VBA storage
 * and MS-OFORMS gives them form-specific child streams.
 */
function classifyVba(streams, storages) {
  const vbaPrefix = 'Root Entry/VBA/';
  const specialVbaStreams = new Set(['_VBA_PROJECT', 'dir', '_dir']);
  const moduleStreamNames = [...streams.keys()]
    .filter((path) => path.startsWith(vbaPrefix))
    .map((path) => path.slice(vbaPrefix.length))
    .filter((name) => !name.includes('/'))
    .filter((name) => !specialVbaStreams.has(name))
    .filter((name) => !/^__SRP_/i.test(name))
    .sort();

  const designerStoragePaths = storages.filter((storage) => {
    const prefix = storage.endsWith('/') ? storage : `${storage}/`;
    if (prefix === vbaPrefix || prefix === 'Root Entry/') {
      return false;
    }
    const directChildren = [...streams.keys()]
      .filter((path) => path.startsWith(prefix))
      .map((path) => path.slice(prefix.length))
      .filter((name) => name && !name.includes('/'))
      .map((name) => name.replace(/^\u0003/, '').toLowerCase());
    return (
      directChildren.includes('vbframe') ||
      directChildren.includes('f') ||
      directChildren.includes('o')
    );
  });
  const designerStorageNames = designerStoragePaths
    .map((path) => path.replace(/\/$/, '').split('/').at(-1))
    .sort();
  const sheetModuleCandidates = moduleStreamNames.filter((name) =>
    /^(Feuil|Sheet|Blatt|Hoja|Foglio)\d+$/i.test(name)
  );
  const otherModuleStreams = moduleStreamNames.filter(
    (name) =>
      name !== 'ThisWorkbook' &&
      !sheetModuleCandidates.includes(name) &&
      !designerStorageNames.includes(name)
  );

  return {
    hasVba: streams.size > 0,
    streamCount: streams.size,
    storageCount: storages.length,
    moduleStreamCount: moduleStreamNames.length,
    moduleStreamNames,
    sheetModuleCandidates,
    otherModuleStreams,
    hasThisWorkbookStream: moduleStreamNames.includes('ThisWorkbook'),
    designerStorageCount: designerStorageNames.length,
    designerStorageNames,
    userFormCount: designerStorageNames.length,
    userFormNames: designerStorageNames,
  };
}

const [beforePath, afterPath] = process.argv.slice(2);
if (!beforePath || !afterPath) {
  throw new Error(
    'Usage: node test/compare-vba-project-streams.mjs <before.bin> <after.bin>'
  );
}

const before = readCfb(beforePath);
const after = readCfb(afterPath);

const allPaths = [
  ...new Set([...before.streams.keys(), ...after.streams.keys()]),
].sort();
const missing = allPaths.filter(p => before.streams.has(p) && !after.streams.has(p));
const added = allPaths.filter(p => !before.streams.has(p) && after.streams.has(p));
const changed = allPaths.filter(p => {
  const prev = before.streams.get(p);
  const curr = after.streams.get(p);
  return (
    prev && curr &&
    (prev.bytes !== curr.bytes || prev.sha256 !== curr.sha256)
  );
});
const allStoragePaths = [
  ...new Set([...before.storages, ...after.storages]),
].sort();
const missingStorages = allStoragePaths.filter(
  (path) => before.storages.includes(path) && !after.storages.includes(path)
);
const addedStorages = allStoragePaths.filter(
  (path) => !before.storages.includes(path) && after.storages.includes(path)
);

const result = {
  ok:
    missing.length === 0 &&
    added.length === 0 &&
    changed.length === 0 &&
    missingStorages.length === 0 &&
    addedStorages.length === 0,
  beforeStreamCount: before.streams.size,
  afterStreamCount: after.streams.size,
  beforeStorageCount: before.storages.length,
  afterStorageCount: after.storages.length,
  before: {
    streamCount: before.streams.size,
    storageCount: before.storages.length,
    ...classifyVba(before.streams, before.storages),
  },
  after: {
    streamCount: after.streams.size,
    storageCount: after.storages.length,
    ...classifyVba(after.streams, after.storages),
  },
  missing,
  added,
  changed,
  missingStorages,
  addedStorages,
};

console.log(JSON.stringify(result, null, 2));
if (!result.ok) {
  process.exitCode = 1;
}
