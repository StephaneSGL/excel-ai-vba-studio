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

console.log(`GitHub Actions workflow validation passed: ${workflowNames.length} file(s).`);
