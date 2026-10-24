import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const workflowDirectory = path.join(root, '.github', 'workflows');
const workflowNames = (await readdir(workflowDirectory))
  .filter(name => /\.ya?ml$/i.test(name))
  .sort();

assert.ok(workflowNames.length > 0, 'at least one GitHub Actions workflow is required');

for (const workflowName of workflowNames) {
  const source = await readFile(path.join(workflowDirectory, workflowName), 'utf8');
  assert.doesNotMatch(
    source,
    /^\s*pull_request_target\s*:/m,
    `${workflowName} must not run untrusted pull-request code with target-repository privileges`,
  );
  assert.match(
    source,
    /^permissions:\s*$/m,
    `${workflowName} must declare explicit workflow permissions`,
  );

  for (const match of source.matchAll(/^\s*uses:\s*([^\s#]+)(?:\s+#.*)?$/gm)) {
    const reference = match[1];
    if (reference.startsWith('./')) {
      continue;
    }
    const separator = reference.lastIndexOf('@');
    assert.ok(separator > 0, `${workflowName} contains an action without a ref: ${reference}`);
    const revision = reference.slice(separator + 1);
    assert.match(
      revision,
      /^[0-9a-f]{40}$/,
      `${workflowName} must pin ${reference.slice(0, separator)} to an immutable commit SHA`,
    );
  }
}

const main = await readFile(path.join(workflowDirectory, 'main.yml'), 'utf8');
assert.match(main, /extension-host:/, 'installed-VSIX integration must be a required validation job');
assert.match(main, /target: linux-x64/, 'VSIX smoke matrix must cover Linux');
assert.match(main, /target: win32-x64/, 'VSIX smoke matrix must cover Windows');
assert.match(main, /xvfb-run -a npm run test:extension-host/, 'Linux must render the custom editor under Xvfb');
assert.match(main, /run: npm run test:automation-safety/, 'Windows must execute the COM policy tests');
const publishing = await readFile(path.join(workflowDirectory, 'publish.yml'), 'utf8');
assert.match(publishing, /target: \[win32-x64, linux-x64\]/, 'release must build both qualified platform targets');
assert.match(publishing, /npm run package -- --target \$\{\{ matrix\.target \}\}/, 'release packaging must bind the artifact target');

console.log(`GitHub Actions workflow validation passed: ${workflowNames.length} file(s).`);
